package com.example.timesheet;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.List;
import java.util.Map;

public class ExcelWriter {
    private static final String FILE_PATH = "timesheet.xlsx";

    public static void write(TimesheetRequest request) throws IOException {
        File file = new File(FILE_PATH);
        Workbook workbook;
        Sheet sheet;

        // Create or load workbook
        if (file.exists()) {
            try (InputStream in = new FileInputStream(file)) {
                workbook = WorkbookFactory.create(in);
                sheet = workbook.getSheetAt(0);
            }
        } else {
            workbook = new XSSFWorkbook();
            sheet = workbook.createSheet("Timesheet");
            createHeaderRow(sheet);
        }

        // Create cell styles
        CellStyle currencyStyle = workbook.createCellStyle();
        DataFormat format = workbook.createDataFormat();
        currencyStyle.setDataFormat(format.getFormat("#,##0.00 €"));

        CellStyle numberStyle = workbook.createCellStyle();
        numberStyle.setDataFormat(format.getFormat("0.0"));

        // Find the next empty row
        int rowNum = sheet.getLastRowNum() + 1;

        // Write data rows
        for (Map<String, String> rowData : request.getTable()) {
            Row row = sheet.createRow(rowNum++);
            int cellIndex = 0;

            // Add department and employee
            row.createCell(cellIndex++).setCellValue(request.getDepartment());
            row.createCell(cellIndex++).setCellValue(request.getEmployee());

            // Add project name
            row.createCell(cellIndex++).setCellValue(rowData.get("Project"));

            // Add month
            row.createCell(cellIndex++).setCellValue(rowData.get("Month"));

            // Add hours with number format
            Cell hoursCell = row.createCell(cellIndex++);
            try {
                double hours = Double.parseDouble(rowData.get("Hours"));
                hoursCell.setCellValue(hours);
                hoursCell.setCellStyle(numberStyle);
            } catch (Exception e) {
                hoursCell.setCellValue(rowData.get("Hours"));
            }

            // Add cost with currency format
            Cell costCell = row.createCell(cellIndex++);
            try {
                double cost = Double.parseDouble(rowData.get("Cost").replace("€", "").replace(",", "").trim());
                costCell.setCellValue(cost);
                costCell.setCellStyle(currencyStyle);
            } catch (Exception e) {
                costCell.setCellValue(rowData.get("Cost"));
            }

            // Add budget with currency format
            Cell budgetCell = row.createCell(cellIndex++);
            try {
                double budget = Double.parseDouble(rowData.get("Budget").replace("€", "").replace(",", "").trim());
                budgetCell.setCellValue(budget);
                budgetCell.setCellStyle(currencyStyle);
            } catch (Exception e) {
                budgetCell.setCellValue(rowData.get("Budget"));
            }

            // Add remaining with currency format
            Cell remainingCell = row.createCell(cellIndex++);
            try {
                double remaining = Double.parseDouble(rowData.get("Remaining").replace("€", "").replace(",", "").trim());
                remainingCell.setCellValue(remaining);
                remainingCell.setCellStyle(currencyStyle);
            } catch (Exception e) {
                remainingCell.setCellValue(rowData.get("Remaining"));
            }
        }

        // Auto-size columns
        for (int i = 0; i < 8; i++) {
            sheet.autoSizeColumn(i);
        }

        // Save workbook
        try (FileOutputStream out = new FileOutputStream(file)) {
            workbook.write(out);
        }
        workbook.close();
    }

    public static synchronized void appendData(Map<String, String> data) throws IOException {
        Workbook workbook;
        Sheet sheet;

        File file = new File(FILE_PATH);

        // Always use try-with-resources to avoid partial writes
        if (file.exists()) {
            try (InputStream inp = new FileInputStream(file)) {
                workbook = WorkbookFactory.create(inp);
            } catch (Exception e) {
                // Corrupt file: delete and start fresh
                file.delete();
                workbook = new XSSFWorkbook();
                sheet = workbook.createSheet("Timesheet");
                createHeaderRow(sheet, data);
            }

            sheet = workbook.getSheetAt(0);
        } else {
            workbook = new XSSFWorkbook();
            sheet = workbook.createSheet("Timesheet");
            createHeaderRow(sheet, data);
        }

        // Append row
        int rowNum = sheet.getLastRowNum() + 1;
        Row row = sheet.createRow(rowNum);

        int i = 0;
        for (String key : data.keySet()) {
            row.createCell(i++).setCellValue(data.get(key));
        }

        // Safely write and close
        try (FileOutputStream out = new FileOutputStream(FILE_PATH)) {
            workbook.write(out);
        }

        workbook.close();
    }

    private static void createHeaderRow(Sheet sheet) {
        Row header = sheet.createRow(0);
        int cellIndex = 0;

        // Add headers
        header.createCell(cellIndex++).setCellValue("Department");
        header.createCell(cellIndex++).setCellValue("Employee");
        header.createCell(cellIndex++).setCellValue("Project");
        header.createCell(cellIndex++).setCellValue("Month");
        header.createCell(cellIndex++).setCellValue("Hours");
        header.createCell(cellIndex++).setCellValue("Cost");
        header.createCell(cellIndex++).setCellValue("Budget");
        header.createCell(cellIndex++).setCellValue("Remaining");
    }

    private static void createHeaderRow(Sheet sheet, Map<String, String> data) {
        Row header = sheet.createRow(0);
        int i = 0;
        for (String key : data.keySet()) {
            header.createCell(i++).setCellValue(key);
        }
    }

    private static void createHeaderRow(Sheet sheet, TimesheetRequest request) {
        Row header = sheet.createRow(0);
        int cellIndex = 0;

        // Add basic headers
        header.createCell(cellIndex++).setCellValue("Department");
        header.createCell(cellIndex++).setCellValue("Employee");
        header.createCell(cellIndex++).setCellValue("Project");
        header.createCell(cellIndex++).setCellValue("Budget");
        header.createCell(cellIndex++).setCellValue("Remaining");

        // Add month-specific headers from the first row of data
        if (!request.getTable().isEmpty()) {
            Map<String, String> firstRow = request.getTable().get(0);
            for (String key : firstRow.keySet()) {
                if (key.contains("Hours")) {
                    header.createCell(cellIndex++).setCellValue(key);
                    header.createCell(cellIndex++).setCellValue(key.replace("Hours", "Cost"));
                }
            }
        }
    }
}
