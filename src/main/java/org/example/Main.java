package org.example;

import Matcher.Excel.Constant;
import Matcher.Excel.ExcelOps;

import java.io.IOException;
import java.text.ParseException;

// Press Shift twice to open the Search Everywhere dialog and type `show whitespaces`,
// then press Enter. You can now see whitespace characters in your code.
public class Main {
    public static void main(String[] args) throws IOException, ParseException {
        // Press Alt+Enter with your caret at the highlighted text to see how
        // IntelliJ IDEA suggests fixing it.
        System.out.println("Hello and welcome!");
        System.out.println("Lets compare the sheets...............");
        ExcelOps.getDatathroughExcel(Constant.SHEET_NAME,Constant.Excel1,Constant.Excel2);
    }
}