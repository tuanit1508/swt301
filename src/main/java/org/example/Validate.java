package org.example;

public class Validate {
    public double checkValid(String value, String variableName, int row) {
        try {
            double number = Double.parseDouble(value);
            if (number < 0 || number > 65535) {
                throw new NumberFormatException(variableName + " out of range (0-65535) in row " + (row + 1) + ": " + number);
            }
            return number;
        } catch (NumberFormatException e) {
            throw new NumberFormatException("Invalid input for " + variableName + " in row " + (row + 1) + ": " + value + ". Please enter a valid number between 0 and 65535.");
        }
    }
}
