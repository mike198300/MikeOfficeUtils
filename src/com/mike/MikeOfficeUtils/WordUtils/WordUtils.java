package com.mike.MikeOfficeUtils.WordUtils;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

public class WordUtils {
    public static String extractDocPlainText(String inputDocPathString) throws IOException {
        String resuleString="";
        WordExtractor docExtractor = new WordExtractor(new FileInputStream(inputDocPathString));
        resuleString = docExtractor.getText();
        return resuleString;
    }

    public static String extractDocxPlainText(String inputDocxPathString) throws IOException {
        String resuleString="";
        XWPFDocument docxDocument = new XWPFDocument(new FileInputStream(inputDocxPathString));
        XWPFWordExtractor docxExtractor = new XWPFWordExtractor(docxDocument);
        resuleString = docxExtractor.getText();
        return resuleString;
    }

    public static void main(String[] args) throws IOException {
//        String text = extractDocPlainText("C:\\Users\\Mike\\Documents\\LIbs\\solr\\test.doc");
        String text = extractDocxPlainText("C:\\Users\\Mike\\Documents\\LIbs\\solr\\test.docx");
        System.out.println(text);
    }
}
