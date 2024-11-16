package com.example;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class MyService {

    public static void performTask() {
        String filePath = "C:\\Users\\satis\\OneDrive\\Desktop\\MSR\\Framework\\spring-apace-poi\\src\\main\\resources\\MSR-Bank-Statement-Oct1-Oct31.xls";
        try (FileInputStream fis = new FileInputStream(filePath);
             HSSFWorkbook workbook = new HSSFWorkbook(fis);) {
             //XSSFWorkbook workbook = new XSSFWorkbook(fis);) {

            Sheet sheet = workbook.getSheetAt(0); // Get the first sheet
            for (Row row : sheet) {
                Cell cell = row.getCell(1); // Get the cell in the second column (index 1)
                if (cell == null) {
                    //cell = row.createCell(1);
                } else {
                    String cellData = cell.getStringCellValue();
                    String lastData = getLastSegmentAfterHyphen(cellData);

                    cell.setCellValue(lastData); // Set new value
                }
            }

            try (FileOutputStream fos = new FileOutputStream(filePath)) {
                workbook.write(fos); // Write changes to the file
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static String getLastSegmentAfterHyphen(String str) {
        int lastHyphenIndex = str.indexOf("-MSR");
        if (lastHyphenIndex == -1) {
            return str; // No hyphen found, return the whole string
        }
        return str.substring(lastHyphenIndex + 4).trim();
    }
}

