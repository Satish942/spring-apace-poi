package com.example;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class DataCategorizer {
    public static Map<String, List<String>> categorizeData(Sheet sheet) {
        Map<String, List<String>> categorizedData = new HashMap<>();

        for (Row row : sheet) {
            String category = row.getCell(0).getStringCellValue();
            String value = row.getCell(1).getStringCellValue();

            categorizedData.computeIfAbsent(category, k -> new ArrayList<>()).add(value);
        }

        return categorizedData;
    }
}
