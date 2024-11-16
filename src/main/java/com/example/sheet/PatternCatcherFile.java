package com.example.sheet;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.regex.Pattern;

public class PatternCatcherFile {
    public static void main(String[] args) {
        String inputFilePath = "input.xlsx";
        String outputFilePath = "filtered_output.xlsx";
        String filterColumn = "Account"; // The column to apply the filter on
        String pattern = ".*Revenue.*";  // The regex pattern to filter by

        try (FileInputStream fis = new FileInputStream(inputFilePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);
            Iterator<Row> rowIterator = sheet.iterator();

            // Identify the column index for filtering
            Row headerRow = rowIterator.next();
            int filterColumnIndex = -1;
            for (Cell cell : headerRow) {
                if (cell.getStringCellValue().equalsIgnoreCase(filterColumn)) {
                    filterColumnIndex = cell.getColumnIndex();
                    break;
                }
            }

            if (filterColumnIndex == -1) {
                System.out.println("Filter column not found.");
                return;
            }

            Workbook outputWorkbook = new XSSFWorkbook();
            Sheet outputSheet = outputWorkbook.createSheet("FilteredData");
            int outputRowIndex = 0;

            // Copy the header row
            Row outputHeaderRow = outputSheet.createRow(outputRowIndex++);
            copyRow(headerRow, outputHeaderRow);

            // Filter rows based on the pattern
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                Cell cell = row.getCell(filterColumnIndex);
                if (cell != null && cell.getCellType() == CellType.STRING && Pattern.matches(pattern, cell.getStringCellValue())) {
                    Row outputRow = outputSheet.createRow(outputRowIndex++);
                    copyRow(row, outputRow);
                }
            }

            try (FileOutputStream fos = new FileOutputStream(outputFilePath)) {
                outputWorkbook.write(fos);
            }
            outputWorkbook.close();
            workbook.close();

            System.out.println("Filtered data has been written to " + outputFilePath);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void copyRow(Row sourceRow, Row targetRow) {
        for (int i = 0; i < sourceRow.getLastCellNum(); i++) {
            Cell sourceCell = sourceRow.getCell(i);
            if (sourceCell != null) {
                Cell targetCell = targetRow.createCell(i);
                switch (sourceCell.getCellType()) {
                    case STRING:
                        targetCell.setCellValue(sourceCell.getStringCellValue());
                        break;
                    case NUMERIC:
                        targetCell.setCellValue(sourceCell.getNumericCellValue());
                        break;
                    case BOOLEAN:
                        targetCell.setCellValue(sourceCell.getBooleanCellValue());
                        break;
                    case FORMULA:
                        targetCell.setCellFormula(sourceCell.getCellFormula());
                        break;
                    default:
                        break;
                }
            }
        }
    }
}
