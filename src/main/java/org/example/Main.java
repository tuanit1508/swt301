package org.example;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class Main {
    public static void main(String[] args) {
        FileInputStream file = null;
        Workbook inputWorkbook = null;
        Workbook outputWorkbook = new XSSFWorkbook();
        Validate validate = new Validate();

        try {
            file = new FileInputStream(new File("D:\\FPT U\\FU-SU24\\SWT301\\data.xlsx"));
            inputWorkbook = new XSSFWorkbook(file);
            Sheet inputSheet = inputWorkbook.getSheetAt(0);
            DataFormatter dataFormatter = new DataFormatter();

            Sheet outputSheet = outputWorkbook.createSheet("Results");
            Row header = outputSheet.createRow(0);
            header.createCell(0).setCellValue("a");
            header.createCell(1).setCellValue("b");
            header.createCell(2).setCellValue("c");
            header.createCell(3).setCellValue("Delta");
            header.createCell(4).setCellValue("Log Message");
            header.createCell(5).setCellValue("x1");
            header.createCell(6).setCellValue("x2");

            // Loop through rows 2 to 12 (indices 1 to 11)
            for (int i = 1; i <= 11; i++) {
                String aStr = dataFormatter.formatCellValue(inputSheet.getRow(i).getCell(0));
                String bStr = dataFormatter.formatCellValue(inputSheet.getRow(i).getCell(1));
                String cStr = dataFormatter.formatCellValue(inputSheet.getRow(i).getCell(2));

                double a = 0, b = 0, c = 0;
                try {
                    a = validate.checkValid(aStr, "a", i);
                    b = validate.checkValid(bStr, "b", i);
                    c = validate.checkValid(cStr, "c", i);

                    QuadraticEquation equation = new QuadraticEquation(a, b, c);
                    equation.solve(outputSheet, i);
                } catch (InvalidEquationException e) {
                    Row row = outputSheet.createRow(i);
                    row.createCell(0).setCellValue(aStr);
                    row.createCell(1).setCellValue(bStr);
                    row.createCell(2).setCellValue(cStr);
                    row.createCell(4).setCellValue(e.getMessage());
                } catch (NumberFormatException e) {
                    Row row = outputSheet.createRow(i);
                    row.createCell(0).setCellValue(aStr);
                    row.createCell(1).setCellValue(bStr);
                    row.createCell(2).setCellValue(cStr);
                    row.createCell(4).setCellValue("Input invalid");
                }
            }

            try (FileOutputStream outFile = new FileOutputStream(new File("D:\\FPT U\\FU-SU24\\SWT301\\output.xlsx"))) {
                outputWorkbook.write(outFile);
            }

        } catch (IOException e) {
            System.out.println("Error reading the file: " + e.getMessage());
        } finally {
            try {
                if (inputWorkbook != null) {
                    inputWorkbook.close();
                }
                if (file != null) {
                    file.close();
                }
                outputWorkbook.close();
            } catch (IOException e) {
                System.out.println("Error closing the file: " + e.getMessage());
            }
        }
    }
}