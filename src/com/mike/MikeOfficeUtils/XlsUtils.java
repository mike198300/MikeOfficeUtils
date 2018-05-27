package com.mike.MikeOfficeUtils;

// HSSF FOR XLS , XSSF FOR XLSX

import com.alibaba.fastjson.JSON;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.POITextExtractor;
import org.apache.poi.hssf.extractor.ExcelExtractor;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.*;
import java.security.GeneralSecurityException;
import java.util.HashMap;
import java.util.Map;

//import org.apache.poi.hssf.extractor.ExcelExtractor;
//import org.apache.poi.hssf.usermodel.HSSFWorkbook;
public class XlsUtils {

    public static String extractXlsText(String inputXlsPathString) throws IOException {
        InputStream fis = null;
        HSSFWorkbook xlsWB  = null;
        String resString = "";
        try {
            fis = new FileInputStream(inputXlsPathString);
            xlsWB = new HSSFWorkbook(fis);
            ExcelExtractor xlsExtractor = new ExcelExtractor(xlsWB);
            xlsExtractor.setFormulasNotResults(true);
            xlsExtractor.setIncludeSheetNames(true);
            resString = xlsExtractor.getText();
        }catch (Exception ex) {
            ex.printStackTrace();
        } finally {
            try {
                assert (xlsWB != null):"Try to close a not existed workbook";
                xlsWB.close();

                fis.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return resString;
    }

    public static String extractXlsMetaDataJson(String inputXlsPathString) {
        InputStream fis = null;
        HSSFWorkbook xlsWB  = null;
        String resString = "";
        try {
            fis = new FileInputStream(inputXlsPathString);
            xlsWB = new HSSFWorkbook(fis);
            ExcelExtractor xlsExtractor = new ExcelExtractor(xlsWB);
            POITextExtractor metaExtractor = xlsExtractor.getMetadataTextExtractor();
            String metaString = metaExtractor.getText();
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
            resString = JSON.toJSONString(metadataMap);
        } catch (IOException e) {
            e.printStackTrace();
        }
        return resString;
    }

    public static boolean isExcelEncrypted(String inputXlsxPathString) throws IOException, GeneralSecurityException {
        return !validateExcelPassword(inputXlsxPathString,"");
    }

    public static boolean validateExcelPassword(String inputExcelFilenameString, String password) throws IOException {
        boolean result = false;
        if (validateExcelFile(inputExcelFilenameString)){
            Workbook wb = null;
            try {
                wb = WorkbookFactory.create(new File(inputExcelFilenameString),password);
                result = true;
            } catch (InvalidFormatException e) {
                e.printStackTrace();
            } catch (EncryptedDocumentException ex) {
                result = false;
            } finally {
                if (wb != null){
                    wb.close();
                }
                return result;
            }
        } else {
            throw new IOException("File validation failed!");
        }
    }

    public static boolean validateExcelFile(String inputFilenameString) throws FileNotFoundException {
        String suffix = OfficeCommonUtils.getFileExtention(inputFilenameString);
        if (!(suffix.equals("XLS") || (suffix.equals("XLSX")))){
            System.err.println("Not a xls or xlsx file.");
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
        String text;
        boolean isEncrypted = false;
        isEncrypted = isExcelEncrypted("C:\\Users\\mike\\IdeaProjects\\MikeOfficeUtils\\testDocs\\test.xlsx");
        isEncrypted = isExcelEncrypted("C:\\Users\\mike\\IdeaProjects\\MikeOfficeUtils\\testDocs\\test_enc.xlsx");
        isEncrypted = validateExcelPassword("C:\\Users\\mike\\IdeaProjects\\MikeOfficeUtils\\testDocs\\test_enc.xlsx", "123");
        isEncrypted = validateExcelPassword("C:\\Users\\mike\\IdeaProjects\\MikeOfficeUtils\\testDocs\\test_enc.xlsx", "000");
        isEncrypted = isExcelEncrypted("C:\\Users\\mike\\IdeaProjects\\MikeOfficeUtils\\testDocs\\test.xls");
        isEncrypted = isExcelEncrypted("C:\\Users\\mike\\IdeaProjects\\MikeOfficeUtils\\testDocs\\test_enc.xls");
        isEncrypted = validateExcelPassword("C:\\Users\\mike\\IdeaProjects\\MikeOfficeUtils\\testDocs\\test_enc.xlsx", "123");
        isEncrypted = validateExcelPassword("C:\\Users\\mike\\IdeaProjects\\MikeOfficeUtils\\testDocs\\test_enc.xlsx", "000");
        text = extractXlsText("C:\\Users\\mike\\IdeaProjects\\MikeOfficeUtils\\testDocs\\test.xls");
        extractXlsMetaDataJson("C:\\Users\\mike\\IdeaProjects\\MikeOfficeUtils\\testDocs\\test.xls");        System.out.println(text);
    }

}
