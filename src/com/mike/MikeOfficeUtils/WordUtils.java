package com.mike.MikeOfficeUtils;

import com.alibaba.fastjson.JSON;
import org.apache.poi.POIXMLProperties;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.openxml4j.opc.internal.PackagePropertiesPart;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

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

    public static String extractDocMetaDataJson(String inputDocPathString) throws IOException {
        WordExtractor docExtractor = new WordExtractor(new FileInputStream(inputDocPathString));
        String metaString = docExtractor.getMetadataTextExtractor().getText();
        Map metadataMap = new HashMap();
        for (String line :
                metaString.split("\n")) {
            String[] buf = line.split("=");
            if (buf[0].startsWith("PID_")){
                if (buf[1]==null){
                    metadataMap.put(buf[0].trim(), "");
                } else {
                    metadataMap.put(buf[0].trim(), buf[1].trim());
                }
            }
        }
        String res = JSON.toJSONString(metadataMap);
        return res;
    }

    public static boolean isWordFileEncrypted(String inputWordFilenameString){
        return true;
    }

//    public static boolean

    public static String extractDocxMetaDataJson(String inputDocxPathString) throws IOException {
        XWPFDocument docxDocument = new XWPFDocument(new FileInputStream(inputDocxPathString));
        XWPFWordExtractor docxExtractor = new XWPFWordExtractor(docxDocument);
        String aaa;
        POIXMLProperties bbb = docxDocument.getProperties();
        POIXMLProperties.CoreProperties ccc = bbb.getCoreProperties();
        PackagePropertiesPart fff = ccc.getUnderlyingProperties();
        aaa = fff.toString();
//        POIXMLProperties.CustomProperties ddd = bbb.getCustomProperties();
//        POIXMLProperties.ExtendedProperties eee = bbb.getExtendedProperties();
        aaa = docxDocument.getProperties().getCoreProperties().getUnderlyingProperties().toString();
        return "";
    }

    public static boolean validateExcelFile(String inputFilenameString) throws FileNotFoundException {
        String suffix = OfficeCommonUtils.getFileExtention(inputFilenameString);
        if (!(suffix.equals("DOC") || (suffix.equals("DOCX")))){
            System.err.println("Not a doc or docx file.");
            return false;
        }
        if (new File(inputFilenameString).exists()){
            return true;
        } else {
            System.err.println("File does not exist.");
            return false;
        }
    }

    public static void main(String[] args) throws IOException {
//        String text = extractDocPlainText("C:\\Users\\Mike\\Documents\\LIbs\\solr\\test.doc");
        String text;
//        text = extractDocxPlainText("C:\\Users\\Mike\\Documents\\LIbs\\solr\\test.docx");
        text = extractDocMetaDataJson("C:\\Users\\mike\\IdeaProjects\\MikeOfficeUtils\\testDocs\\test.doc");
//        text = extractDocxMetaDataJson("C:\\Users\\mike\\IdeaProjects\\MikeOfficeUtils\\testDocs\\test.docx");
        System.out.println(text);
    }
}
