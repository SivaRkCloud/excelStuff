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
        System.out.println("Lets compare the Files..");
        String file1 = Constant.Excel1 + args[0];
        String file2 = Constant.Excel2 + args[1];

        System.out.println( file1 );
        System.out.println(file2 );

        File f = new File(file1);
        if (!f.exists())
        {
            System.out.println( "File:["+file1 + "] Not found in the path..");
            return;
        }
        f = new File(file2);
        if (!f.exists())
        {
            System.out.println( "File:["+file2 + "] Not found in the path..");
            return;
        }

        String file3=System.getProperty("user.dir")+"\\output\\"+"diff_report.xlsx";
//        File source = new File(Constant.Excel1);
//        File dest = new File(Excel3);
        //copyFile(source,dest);
        String Excel3="diff_report";

        ExcelOps.getDatathroughExcel(Constant.SHEET_NAME,file1,file2,file3);
        
        System.out.println("Files..Compared.. successfulLy..");

    }
    private static void copyFile(File source, File dest) throws IOException {
        Files.copy(source.toPath(), dest.toPath());
    }
}