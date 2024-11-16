package com.example;

import com.example.DataCategorizer;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Map;

public class ExcelWriter {
    public static void writeData(Map<String, List<String>> data, String outputPath) throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Categorized Data");

        int rowNum = 0;
        for (Map.Entry<String, List<String>> entry : data.entrySet()) {
            Row row = sheet.createRow(rowNum++);
            Cell categoryCell = row.createCell(0);
            categoryCell.setCellValue(entry.getKey());

            int cellNum = 1;
            for (String value : entry.getValue()) {
                Cell valueCell = row.createCell(cellNum++);
                valueCell.setCellValue(value);
            }
        }

        FileOutputStream fileOut = new FileOutputStream(outputPath);
        workbook.write(fileOut);
        fileOut.close();
        workbook.close();
    }

    public static void main(String[] args) throws IOException {
        FileInputStream file = new FileInputStream("input.xlsx");
        Workbook workbook = new XSSFWorkbook(file);
        Sheet sheet = workbook.getSheetAt(0);

        Map<String, List<String>> categorizedData = DataCategorizer.categorizeData(sheet);
        writeData(categorizedData, "output.xlsx");

        workbook.close();
        file.close();
    }
}
