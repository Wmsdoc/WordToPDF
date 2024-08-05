package com.example.demo;

import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

@SpringBootTest
class DemoApplicationTests {

    @Test
    void contextLoads() throws IOException {
        File file = new File("C:\\Users\\29322\\Desktop\\1.docx");

        DocumentConverter documentConverter = new DocumentConverter();

        byte[] bytes = documentConverter.convertToPdf(new FileInputStream(file));


        File pdfFile = new File("C:\\Users\\29322\\Desktop\\1.pdf");
        FileOutputStream fileOutputStream = new FileOutputStream(pdfFile);
        fileOutputStream.write(bytes);
        fileOutputStream.close();

    }

}
