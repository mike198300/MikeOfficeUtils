package com.mike.MikeOfficeUtils;

public class OfficeCommonUtils {
    public static String getFileExtention(String inputFilenameString){
        return inputFilenameString.substring(inputFilenameString.lastIndexOf(".") + 1, inputFilenameString.length()).toUpperCase();
    }
}
