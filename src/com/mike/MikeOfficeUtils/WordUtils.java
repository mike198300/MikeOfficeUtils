package com.mike.MikeOfficeUtils;

import com.alibaba.fastjson.JSON;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.POIXMLProperties;
import org.apache.poi.hssf.record.crypto.Biff8EncryptionKey;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.openxml4j.opc.internal.PackagePropertiesPart;
import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.filesystem.NPOIFSFileSystem;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.*;
import java.security.GeneralSecurityException;
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

    public static boolean validateWordPassword(String inputWordFilenameString, String password) throws IOException, GeneralSecurityException {
        boolean result = false;
        if (validateWordFile(inputWordFilenameString)){
            String suffix = OfficeCommonUtils.getFileExtention(inputWordFilenameString);
            switch (suffix){
                case "DOC":
                    HWPFDocument docDocument = null;
                    POIFSFileSystem pfs = null;
                    try {
                        Biff8EncryptionKey.setCurrentUserPassword(password);
                        pfs = new POIFSFileSystem(new FileInputStream(inputWordFilenameString));
                        docDocument=new HWPFDocument(pfs);
                        result = true;
                    } catch (EncryptedDocumentException ex) {
                        result = false;
                    }catch (IOException e) {
                        e.printStackTrace();
                    } finally {
                        if(docDocument != null){
                            docDocument.close();
                        }
                        if (pfs != null){
                            pfs.close();
                        }
                    }
                    break;
                case "DOCX":
                    if (password.equals("")) {
                        XWPFDocument docxDocument = null;
                        InputStream is = null;
                        try {
                            is = new FileInputStream(inputWordFilenameString);
                            docxDocument = new XWPFDocument(is);
                            result = true;
                        } catch (FileNotFoundException e) {
                            e.printStackTrace();
                        } catch (Exception ex) {
                            result = false;
                        }finally {
                            if (is != null){
                                is.close();
                            }
                            if (docxDocument != null){
                                docxDocument.close();
                            }
                        }
                    } else {
                        NPOIFSFileSystem npfs = null;
                        InputStream is = null;
                        try {
                            is = new FileInputStream(inputWordFilenameString);
                            npfs= new NPOIFSFileSystem(is);
                            EncryptionInfo encInfo = new EncryptionInfo(npfs);
                            Decryptor d = encInfo.getDecryptor();
                            result = d.verifyPassword(password);
                        } catch (IOException e) {
                            e.printStackTrace();
                        }  finally {
                            if (is != null){
                                is.close();
                            }
                            if (npfs != null){
                                npfs.close();
                            }
                        }
                    }
                    break;
            }
            return result;
        } else {
            throw new IOException("File validation failed!");
        }
    }

    public static boolean validateWordFile(String inputFilenameString) throws FileNotFoundException {
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

    public static void main(String[] args) throws IOException, GeneralSecurityException {
//        String text = extractDocPlainText("C:\\Users\\Mike\\Documents\\LIbs\\solr\\test.doc");
//        String text;
//        text = extractDocxPlainText("C:\\Users\\Mike\\Documents\\LIbs\\solr\\test.docx");
//        text = extractDocMetaDataJson("C:\\Users\\mike\\IdeaProjects\\MikeOfficeUtils\\testDocs\\test.doc");
//        text = extractDocxMetaDataJson("C:\\Users\\mike\\IdeaProjects\\MikeOfficeUtils\\testDocs\\test.docx");
        boolean result = false;
        result = validateWordPassword("C:\\Users\\mike\\IdeaProjects\\MikeOfficeUtils\\testDocs\\test.doc", "");
        result = validateWordPassword("C:\\Users\\mike\\IdeaProjects\\MikeOfficeUtils\\testDocs\\test_enc.doc", "");
        result = validateWordPassword("C:\\Users\\mike\\IdeaProjects\\MikeOfficeUtils\\testDocs\\test_enc.doc", "123");
        result = validateWordPassword("C:\\Users\\mike\\IdeaProjects\\MikeOfficeUtils\\testDocs\\test.docx", "");
        result = validateWordPassword("C:\\Users\\mike\\IdeaProjects\\MikeOfficeUtils\\testDocs\\test_enc.docx", "");
        result = validateWordPassword("C:\\Users\\mike\\IdeaProjects\\MikeOfficeUtils\\testDocs\\test_enc.docx", "123");
        result = validateWordPassword("C:\\Users\\mike\\IdeaProjects\\MikeOfficeUtils\\testDocs\\test_enc.docx", "456");
        System.out.println(result);
    }
}
