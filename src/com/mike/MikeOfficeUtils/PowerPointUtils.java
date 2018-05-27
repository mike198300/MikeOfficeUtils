package com.mike.MikeOfficeUtils;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hslf.extractor.PowerPointExtractor;
import org.apache.poi.hslf.usermodel.HSLFSlideShow;
import org.apache.poi.hssf.record.crypto.Biff8EncryptionKey;
import org.apache.poi.openxml4j.exceptions.OLE2NotOfficeXmlFileException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.filesystem.NPOIFSFileSystem;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xslf.extractor.XSLFPowerPointExtractor;
import org.apache.poi.xslf.usermodel.XSLFSlideShow;
import org.apache.xmlbeans.XmlException;

import java.io.*;
import java.security.GeneralSecurityException;

public class PowerPointUtils {

    public static String extractPptPlainText(String inputFilenameString) throws IOException {
        PowerPointExtractor pptExtractor = new PowerPointExtractor(inputFilenameString);
        return pptExtractor.getText();
    }

    public static String extractPptxPlainText(String inputFilenameString) throws IOException, OpenXML4JException, XmlException {
        XSLFSlideShow pptxDocument = new XSLFSlideShow(inputFilenameString);
        XSLFPowerPointExtractor pptxExtracor = new XSLFPowerPointExtractor(pptxDocument);
        return pptxExtracor.getText();
    }

    public static boolean validatePowerPointPassword(String inputFilenameString, String password) throws IOException, GeneralSecurityException, XmlException {
        boolean result = false;
        if (validatePowerPointFile(inputFilenameString)){
            String suffix = OfficeCommonUtils.getFileExtention(inputFilenameString);
            switch (suffix){
                case "PPT":
                    HSLFSlideShow pptDocument = null;
                    NPOIFSFileSystem npfs = null;
                    InputStream is = null;
                    try {
                        Biff8EncryptionKey.setCurrentUserPassword(password);
                        is = new FileInputStream(inputFilenameString);
                        npfs = new POIFSFileSystem(is);
                        pptDocument=new HSLFSlideShow(npfs);
                        result = true;
                    } catch (EncryptedDocumentException ex) {
                        result = false;
                    }catch (IOException e) {
                        e.printStackTrace();
                    } finally {
                        if(pptDocument != null){
                            pptDocument.close();
                        }
                        if (npfs != null){
                            npfs.close();
                        }
                        if (is != null){
                            is.close();
                        }
                    }
                    break;
                case "PPTX":
                    if (password.equals("")) {
                        XSLFSlideShow pptxDocument = null;
//                        InputStream pptxInputStream = null;
                        try {
//                            pptxInputStream = new FileInputStream(inputFilenameString);
                            pptxDocument = new XSLFSlideShow(inputFilenameString);
                            result = true;
                        } catch (FileNotFoundException | OpenXML4JException e) {
                            e.printStackTrace();
                        } catch (OLE2NotOfficeXmlFileException ex) {
                            result = false;
                        } finally {
//                            if (is != null){
//                                is.close();
//                            }
                            if (pptxDocument != null){
                                pptxDocument.close();
                            }
                        }
                    } else {
                        NPOIFSFileSystem pptxNpfs = null;
                        InputStream pptxInputStream = null;
                        try {
                            pptxInputStream = new FileInputStream(inputFilenameString);
                            pptxNpfs= new NPOIFSFileSystem(pptxInputStream);
                            EncryptionInfo encInfo = new EncryptionInfo(pptxNpfs);
                            Decryptor d = encInfo.getDecryptor();
                            result = d.verifyPassword(password);
                        } catch (IOException e) {
                            e.printStackTrace();
                        }  finally {
                            if (pptxNpfs != null){
                                pptxNpfs.close();
                            }
                            if (pptxInputStream != null){
                                pptxInputStream.close();
                            }
                        }
                    }
                    break;
            }
            return result;
        }else {
            throw new IOException("File validation failed!");
        }
    }

    public static boolean validatePowerPointFile(String inputFilenameString) throws FileNotFoundException {
        String suffix = OfficeCommonUtils.getFileExtention(inputFilenameString);
        if (!(suffix.equals("PPT") || (suffix.equals("PPTX")))){
            System.err.println("Not a ppt or pptx file.");
            return false;
        }
        if (new File(inputFilenameString).exists()){
            return true;
        } else {
            System.err.println("File does not exist.");
            return false;
        }
    }

    public static void main(String[] args) throws IOException, GeneralSecurityException, XmlException, OpenXML4JException {

        String text = "";
        text = extractPptPlainText("C:\\Users\\mike\\IdeaProjects\\MikeOfficeUtils\\testDocs\\test.ppt");
        text = extractPptxPlainText("C:\\Users\\mike\\IdeaProjects\\MikeOfficeUtils\\testDocs\\test.pptx");
        boolean result = false;
        result = validatePowerPointPassword("C:\\Users\\mike\\IdeaProjects\\MikeOfficeUtils\\testDocs\\test.ppt", "");
        result = validatePowerPointPassword("C:\\Users\\mike\\IdeaProjects\\MikeOfficeUtils\\testDocs\\test_enc.ppt", "");
        result = validatePowerPointPassword("C:\\Users\\mike\\IdeaProjects\\MikeOfficeUtils\\testDocs\\test_enc.ppt", "123");
        result = validatePowerPointPassword("C:\\Users\\mike\\IdeaProjects\\MikeOfficeUtils\\testDocs\\test.pptx", "");
        result = validatePowerPointPassword("C:\\Users\\mike\\IdeaProjects\\MikeOfficeUtils\\testDocs\\test_enc.pptx", "");
        result = validatePowerPointPassword("C:\\Users\\mike\\IdeaProjects\\MikeOfficeUtils\\testDocs\\test_enc.pptx", "123");
        result = validatePowerPointPassword("C:\\Users\\mike\\IdeaProjects\\MikeOfficeUtils\\testDocs\\test_enc.pptx", "456");
    }
}
