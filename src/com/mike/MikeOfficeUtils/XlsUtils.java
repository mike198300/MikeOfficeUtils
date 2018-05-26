package com.mike.MikeOfficeUtils;

// HSSF FOR XLS , XSSF FOR XLSX

import com.alibaba.fastjson.JSON;
import org.apache.poi.POITextExtractor;
import org.apache.poi.hssf.extractor.ExcelExtractor;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;

//import org.apache.poi.hssf.extractor.ExcelExtractor;
//import org.apache.poi.hssf.usermodel.HSSFWorkbook;
public class XlsUtils {

    public static String extractXlsText(String inputDocPathString) throws IOException {
        InputStream fis = null;
        HSSFWorkbook xlsWB  = null;
        String resString = "";
        try {
            fis = new FileInputStream(inputDocPathString);
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

    public static String extractXlsMetaDataJson(String inputDocPathString) {
        InputStream fis = null;
        HSSFWorkbook xlsWB  = null;
        String resString = "";
        try {
            fis = new FileInputStream(inputDocPathString);
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

    public static void main(String[] args) throws IOException {
        String text;
        text = extractXlsText("C:\\Users\\mike\\IdeaProjects\\MikeOfficeUtils\\testDocs\\test.xls");
        extractXlsMetaDataJson("C:\\Users\\mike\\IdeaProjects\\MikeOfficeUtils\\testDocs\\test.xls");        System.out.println(text);
    }

}
