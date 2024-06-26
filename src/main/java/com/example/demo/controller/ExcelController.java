package com.example.demo.controller;

import org.apache.poi.ss.usermodel.*;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

@RestController
public class ExcelController {

    @GetMapping("/read-excel")
    public String readExcel() {
        String excelFilePath = "C:/Users/Renu/Documents/asignmentexcel.xlsx"; 
        StringBuilder data = new StringBuilder();

        try (FileInputStream fis = new FileInputStream(new File(excelFilePath));
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);
            for (Row row : sheet) {
                for (Cell cell : row) {
                    switch (cell.getCellType()) {
                        case STRING:
                            data.append(cell.getStringCellValue()).append("\t");
                            break;
                        case NUMERIC:
                            data.append(cell.getNumericCellValue()).append("\t");
                            break;
                        default:
                            data.append("UNKNOWN\t");
                            break;
                    }
                }
                data.append("\n");
            }

        } catch (IOException e) {
            e.printStackTrace();
            return "Error reading the Excel file.";
        }

        return data.toString();
    }
}