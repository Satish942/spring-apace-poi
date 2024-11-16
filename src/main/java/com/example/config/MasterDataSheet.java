package com.example.config;

import org.apache.poi.ss.usermodel.*;
import org.springframework.stereotype.Component;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@Component
public class MasterDataSheet {

    public static Map<String, List<String>> categorizeData(Sheet sheet) {
        Map<String, List<String>> categorizedData = new HashMap<>();

        for (Row row : sheet) {
            String category = row.getCell(0).getStringCellValue();
            String value = row.getCell(1).getStringCellValue();

            categorizedData.computeIfAbsent(category, k -> new ArrayList<>()).add(value);
        }

        return categorizedData;
    }

    public Map<String, List<Object>> readExcel(MultipartFile file) throws IOException {
        Map<String, List<Object>> dataMap = new HashMap<>();
        try (Workbook workbook = WorkbookFactory.create(file.getInputStream())) {
            Sheet sheet = workbook.getSheetAt(0);
            Row headerRow = sheet.getRow(0);

            for (Cell cell : headerRow) {
                dataMap.put(cell.getStringCellValue(), new ArrayList<>());
            }

            for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                for (int cellIndex = 0; cellIndex < headerRow.getPhysicalNumberOfCells(); cellIndex++) {
                    if (row != null) {
                        Cell cell = row.getCell(cellIndex);
                        String key = headerRow.getCell(cellIndex).getStringCellValue();
                        if (cell != null) {
                            dataMap.get(key).add(cell.toString());
                        }
                    }
                }
            }
        }
        return dataMap;
    }

}
