package org.example;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.example.InvalidEquationException;

public class QuadraticEquation {
    private double a, b, c;

    public QuadraticEquation(double a, double b, double c) {
        this.a = a;
        this.b = b;
        this.c = c;
    }

    public void solve(Sheet outputSheet, int rowIndex) throws InvalidEquationException {
        Row row = outputSheet.createRow(rowIndex);
        row.createCell(0).setCellValue(a);
        row.createCell(1).setCellValue(b);
        row.createCell(2).setCellValue(c);

        if (a == 0 && b != 0) {

            throw new InvalidEquationException("exception");
        }

        double delta = b * b - 4 * a * c;

        if (delta < 0) {
            row.createCell(3).setCellValue("delta < 0");
            row.createCell(4).setCellValue("The equation has no solution");
        } else if (delta == 0) {
            double x = -b / (2 * a);
            row.createCell(3).setCellValue("delta = 0");
            row.createCell(5).setCellValue(x);
        } else {
            double x1 = (-b + Math.sqrt(delta)) / (2 * a);
            double x2 = (-b - Math.sqrt(delta)) / (2 * a);
            row.createCell(3).setCellValue("delta > 0");
            row.createCell(5).setCellValue(x1);
            row.createCell(6).setCellValue(x2);
        }
    }
}
