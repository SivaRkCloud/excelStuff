package org.example;

import Matcher.Excel.Constant;
import Matcher.Excel.ExcelOps;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.text.ParseException;

// Press Shift twice to open the Search Everywhere dialog and type `show whitespaces`,
// then press Enter. You can now see whitespace characters in your code.
public class Main {
    public static void main(String[] args) throws IOException, ParseException {
        // Press Alt+Enter with your caret at the highlighted text to see how
        // IntelliJ IDEA suggests fixing it.
        System.out.println("Hello and welcome!");
        System.out.println("Lets compare the sheets...............");

//        String Excel3=System.getProperty("user.dir")+"\\output\\"+"diff_report.xlsx";
//        File source = new File(Constant.Excel1);
//        File dest = new File(Excel3);
        //copyFile(source,dest);
        String Excel3="diff_report";

        ExcelOps.getDatathroughExcel(Constant.SHEET_NAME,Constant.Excel1,Constant.Excel2,Excel3);

    }
    private static void copyFile(File source, File dest) throws IOException {
        Files.copy(source.toPath(), dest.toPath());
    }
}