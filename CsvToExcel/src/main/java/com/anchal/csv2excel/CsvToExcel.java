package com.anchal.csv2excel;

import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
 
import org.apache.commons.lang.math.NumberUtils;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import com.opencsv.CSVParser;
import com.opencsv.CSVParserBuilder;
import com.opencsv.CSVReader;
import com.opencsv.CSVReaderBuilder;
 
public class CsvToExcel {
 
    public static final char FILE_DELIMITER = ',';
    public static final String FILE_EXTN = ".xlsx";
    public static final String FILE_NAME = "EXCEL_DATA";
 
 
    public static String convertCsvToXlsx(String xlsFileLocation, String csvFilePath) {
        SXSSFSheet sheet = null;
        CSVReader reader = null;
        Workbook workBook = null;
        CSVParser csvParser = null;
        String generatedXlsFilePath = "";
        FileOutputStream fileOutputStream = null;
 
        try {
 
            /**** Get the CSVReader Instance & Specify The Delimiter To Be Used ****/
            String[] nextLine;
    
            csvParser = new CSVParserBuilder().withSeparator(FILE_DELIMITER).build();

            reader = new CSVReaderBuilder(new FileReader(csvFilePath)).withCSVParser(csvParser).build();
            
            workBook = new SXSSFWorkbook();
            sheet = (SXSSFSheet) workBook.createSheet("Sheet");
 
            int rowNum = 0;
            System.out.println("Creating New .Xlsx File From The Already Generated .Csv File");
            while((nextLine = reader.readNext()) != null) {
                Row currentRow = sheet.createRow(rowNum++);
                for(int i=0; i < nextLine.length; i++) {
                    if(NumberUtils.isDigits(nextLine[i])) {
                        currentRow.createCell(i).setCellValue(Integer.parseInt(nextLine[i]));
                    } else if (NumberUtils.isNumber(nextLine[i])) {
                        currentRow.createCell(i).setCellValue(Double.parseDouble(nextLine[i]));
                    } else {
                        currentRow.createCell(i).setCellValue(nextLine[i]);
                    }
                }
            }
 
            generatedXlsFilePath = xlsFileLocation + FILE_NAME + FILE_EXTN;
            System.out.println("The File Is Generated At The Following Location?= " + generatedXlsFilePath);
 
            fileOutputStream = new FileOutputStream(generatedXlsFilePath.trim());
            workBook.write(fileOutputStream);
        } catch(Exception exObj) {
        	System.out.println("Error: Exception In convertCsvToXls() Method?=  " + exObj);
        } finally {         
            try {
 
                /**** Closing The Excel Workbook Object ****/
                workBook.close();
 
                /**** Closing The File-Writer Object ****/
                fileOutputStream.close();
 
                /**** Closing The CSV File-ReaderObject ****/
                reader.close();
            } catch (IOException ioExObj) {
            	System.out.println("Error: Exception While Closing I/O Objects In convertCsvToXls() Method?=  " + ioExObj);          
            }
        }
 
        return generatedXlsFilePath;
    }   
}