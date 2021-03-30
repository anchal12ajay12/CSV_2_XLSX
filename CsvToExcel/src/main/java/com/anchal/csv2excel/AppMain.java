package com.anchal.csv2excel;
 
public class AppMain {
 
    public static void main(String[] args) {
 
        String xlsLoc = "config/", csvLoc = "config/sample.csv", fileLoc = "";
        fileLoc = CsvToExcel.convertCsvToXlsx(xlsLoc, csvLoc);
        System.out.println("File Location Is?= " + fileLoc);
    }
}