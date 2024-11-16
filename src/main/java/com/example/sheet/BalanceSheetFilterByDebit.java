package com.example.sheet;

import org.apache.poi.ss.usermodel.*;
        import org.apache.poi.xssf.usermodel.XSSFWorkbook;

        import java.io.FileInputStream;
        import java.io.FileOutputStream;
        import java.io.IOException;
        import java.util.Iterator;

public class BalanceSheetFilterByDebit {
    public static void filterByDebit() {
        String inputFilePath = "balance_sheet.xlsx";
        String outputFilePath = "debit_sheet.xlsx";
        String debitColumn = "Deposit Amt."; // The column to apply the filter on

        try (FileInputStream fis = new FileInputStream(inputFilePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);
            Iterator<Row> rowIterator = sheet.iterator();

            // Identify the column index for filtering
            Row headerRow = rowIterator.next();
            int debitColumnIndex = -1;
            for (Cell cell : headerRow) {
                if (cell.getStringCellValue().equalsIgnoreCase(debitColumn)) {
                    debitColumnIndex = cell.getColumnIndex();
                    break;
                }
            }

            if (debitColumnIndex == -1) {
                System.out.println("Debit column not found.");
                return;
            }

            Workbook outputWorkbook = new XSSFWorkbook();
            Sheet outputSheet = outputWorkbook.createSheet("FilteredData");
            int outputRowIndex = 0;

            // Copy the header row
            Row outputHeaderRow = outputSheet.createRow(outputRowIndex++);
            copyRow(headerRow, outputHeaderRow);

            // Filter rows based on the debit column value (non-zero debit values)
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                Cell cell = row.getCell(debitColumnIndex);
                if (cell != null && cell.getCellType() == CellType.NUMERIC && cell.getNumericCellValue() > 0) {
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

    public static void filterByCredit() {
        String inputFilePath = "balance_sheet.xlsx";
        String outputFilePath = "credit_sheet.xlsx";
        String debitColumn = "Withdrawal Amt."; // The column to apply the filter on

        try (FileInputStream fis = new FileInputStream(inputFilePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);
            Iterator<Row> rowIterator = sheet.iterator();

            // Identify the column index for filtering
            Row headerRow = rowIterator.next();
            int debitColumnIndex = -1;
            for (Cell cell : headerRow) {
                if (cell.getStringCellValue().equalsIgnoreCase(debitColumn)) {
                    debitColumnIndex = cell.getColumnIndex();
                    break;
                }
            }

            if (debitColumnIndex == -1) {
                System.out.println("credit column not found.");
                return;
            }

            Workbook outputWorkbook = new XSSFWorkbook();
            Sheet outputSheet = outputWorkbook.createSheet("FilteredData");
            int outputRowIndex = 0;

            // Copy the header row
            Row outputHeaderRow = outputSheet.createRow(outputRowIndex++);
            copyRow(headerRow, outputHeaderRow);

            // Filter rows based on the debit column value (non-zero debit values)
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                Cell cell = row.getCell(debitColumnIndex);
                if (cell != null && cell.getCellType() == CellType.NUMERIC && cell.getNumericCellValue() > 0) {
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
