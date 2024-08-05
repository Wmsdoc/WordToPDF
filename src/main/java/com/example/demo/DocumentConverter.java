package com.example.demo;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.font.PDType1Font;
import org.apache.pdfbox.pdmodel.font.Standard14Fonts;
import org.apache.pdfbox.pdmodel.graphics.image.PDImageXObject;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Picture;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.xwpf.usermodel.*;

import java.io.*;
import java.util.List;

public class DocumentConverter {

    public byte[] convertToPdf(InputStream wordInputStream) throws IOException {
        ByteArrayOutputStream pdfOutputStream = new ByteArrayOutputStream();
        BufferedInputStream bufferedInputStream = new BufferedInputStream(wordInputStream);
        bufferedInputStream.mark(8); // Mark the current position to reset later

        String fileType = detectFileType(bufferedInputStream);
        if ("doc".equalsIgnoreCase(fileType)) {
            bufferedInputStream.reset(); // Reset the stream to the beginning
            convertDocToPdf(bufferedInputStream, pdfOutputStream);
        } else if ("docx".equalsIgnoreCase(fileType)) {
            bufferedInputStream.reset(); // Reset the stream to the beginning
            convertDocxToPdf(bufferedInputStream, pdfOutputStream);
        } else {
            throw new IllegalArgumentException("Unsupported file type");
        }

        return pdfOutputStream.toByteArray();
    }

    private String detectFileType(InputStream inputStream) throws IOException {
        byte[] header = readFileHeader(inputStream, 8);

        if (startsWith(header, new byte[]{(byte) 0xD0, (byte) 0xCF, (byte) 0x11, (byte) 0xE0, (byte) 0xA1, (byte) 0xB1, (byte) 0x1A, (byte) 0xE1})) {
            return "doc";
        } else if (startsWith(header, new byte[]{(byte) 0x50, (byte) 0x4B, (byte) 0x03, (byte) 0x04})) {
            return "docx";
        } else {
            throw new IOException("Unsupported file type");
        }
    }

    private byte[] readFileHeader(InputStream inputStream, int length) throws IOException {
        byte[] header = new byte[length];
        if (inputStream.read(header) != length) {
            throw new IOException("Unable to read enough bytes from the file");
        }
        inputStream.reset();
        return header;
    }

    private boolean startsWith(byte[] header, byte[] magic) {
        if (header.length < magic.length) {
            return false;
        }
        for (int i = 0; i < magic.length; i++) {
            if (header[i] != magic[i]) {
                return false;
            }
        }
        return true;
    }

    private void convertDocToPdf(InputStream inputStream, ByteArrayOutputStream pdfOutputStream) throws IOException {
        HWPFDocument document = new HWPFDocument(inputStream);
        PDDocument pdfDocument = new PDDocument();
        PDPage page = new PDPage();
        pdfDocument.addPage(page);

        float yPosition = 725; // Keep track of the current y-position
        float margin = 25;
        float fontSize = 10; // Adjust font size
        float leading = 12.5f; // Adjust leading

        PDPageContentStream contentStream = new PDPageContentStream(pdfDocument, page);
        try {
            contentStream.beginText();
            contentStream.setFont(new PDType1Font(Standard14Fonts.FontName.HELVETICA), fontSize);
            contentStream.setLeading(leading);
            contentStream.newLineAtOffset(margin, yPosition); // Start writing text at the top

            Range range = document.getRange();
            String text = range.text();

            String[] lines = text.split("\n");
            for (String line : lines) {
                contentStream.showText(line);
                contentStream.newLine();
                yPosition -= leading; // Update y-position after each line
            }

            contentStream.endText();

            List<Picture> pictures = document.getPicturesTable().getAllPictures();
            for (Picture picture : pictures) {
                byte[] pictureData = picture.getContent();
                PDImageXObject pdImage = PDImageXObject.createFromByteArray(pdfDocument, pictureData, null);

                // Get image width and height
                float imageWidth = pdImage.getWidth();
                float imageHeight = pdImage.getHeight();

                // Scale image if it exceeds page width or height
                float scale = Math.min((page.getMediaBox().getWidth() - 2 * margin) / imageWidth, (page.getMediaBox().getHeight() - 2 * margin) / imageHeight);
                if (scale < 1) {
                    imageWidth *= scale;
                    imageHeight *= scale;
                }

                // Adjust y-position to place image below the last line
                yPosition -= imageHeight + 10; // Add some space between text and image

                contentStream.drawImage(pdImage, margin, yPosition, imageWidth, imageHeight);
                yPosition -= imageHeight + 10; // Update y-position after drawing image
            }
        } finally {
            contentStream.close();
        }

        pdfDocument.save(pdfOutputStream);
        pdfDocument.close();
    }

    private void convertDocxToPdf(InputStream inputStream, ByteArrayOutputStream pdfOutputStream) throws IOException {
        XWPFDocument document = new XWPFDocument(inputStream);
        PDDocument pdfDocument = new PDDocument();
        PDPage page = new PDPage();
        pdfDocument.addPage(page);

        float yPosition = 725;
        float margin = 25;
        float fontSize = 10; // Adjust font size
        float leading = 12.5f; // Adjust leading
        float minImageSpacing = 10; // Minimum spacing between text and image

        PDPageContentStream contentStream = new PDPageContentStream(pdfDocument, page);
        try {
            for (XWPFParagraph paragraph : document.getParagraphs()) {
                for (XWPFRun run : paragraph.getRuns()) {
                    // Draw text
                    String text = run.text();
                    if (!text.isEmpty()) {
                        contentStream.beginText();
                        contentStream.setFont(new PDType1Font(Standard14Fonts.FontName.HELVETICA), fontSize);
                        contentStream.setLeading(leading);
                        contentStream.newLineAtOffset(margin, yPosition);
                        contentStream.showText(text);
                        contentStream.endText();
                        yPosition -= leading;
                    }

                    // Draw embedded pictures
                    for (XWPFPicture picture : run.getEmbeddedPictures()) {
                        XWPFPictureData pictureData = picture.getPictureData();
                        PDImageXObject pdImage = PDImageXObject.createFromByteArray(pdfDocument, pictureData.getData(), null);

                        // Get image width and height
                        float imageWidth = pdImage.getWidth();
                        float imageHeight = pdImage.getHeight();

                        // Scale image if it exceeds page width or height
                        float scale = Math.min((page.getMediaBox().getWidth() - 2 * margin) / imageWidth, (page.getMediaBox().getHeight() - 2 * margin) / imageHeight);
                        if (scale < 1) {
                            imageWidth *= scale;
                            imageHeight *= scale;
                        }

                        // Calculate image position
                        float imageYPosition = yPosition - imageHeight - minImageSpacing;

                        // Add new page if image exceeds page height
                        if (imageYPosition < margin) {
                            page = new PDPage();
                            pdfDocument.addPage(page);
                            contentStream.close();
                            contentStream = new PDPageContentStream(pdfDocument, page);
                            yPosition = page.getMediaBox().getHeight() - margin;
                            imageYPosition = yPosition - imageHeight - minImageSpacing;
                        }

                        contentStream.drawImage(pdImage, margin, imageYPosition, imageWidth, imageHeight);
                        yPosition = imageYPosition - minImageSpacing; // Update y-position after drawing image
                    }
                }
            }
        } finally {
            contentStream.close();
        }

        pdfDocument.save(pdfOutputStream);
        pdfDocument.close();
    }

    public static void main(String[] args) {
        DocumentConverter converter = new DocumentConverter();
        try (InputStream wordInputStream = new FileInputStream("C:\\Users\\29322\\Desktop\\1.docx")) {
            byte[] pdfBytes = converter.convertToPdf(wordInputStream);
            try (FileOutputStream fos = new FileOutputStream("C:\\Users\\29322\\Desktop\\1.pdf")) {
                fos.write(pdfBytes);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}