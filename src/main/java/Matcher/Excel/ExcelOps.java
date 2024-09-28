package Matcher.Excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

public class ExcelOps {
    public static boolean searchInTie(Row rowInSheet1, XSSFSheet sheet2) throws IOException {

        Iterator<Row> rowInSheet2 = sheet2.rowIterator();

        // skipping headers
        if (rowInSheet2.hasNext()) {
            rowInSheet2.next();
        }
        // process actual data
        while (rowInSheet2.hasNext()) {
            Row rowIn2 = rowInSheet2.next();
            if (compareDatetime(rowInSheet1, rowIn2) && compareTransit(rowInSheet1, rowIn2) && compareEntityLocation(rowInSheet1, rowIn2) && compareAmount(rowInSheet1, rowIn2))
                return true;
        }

        return false;
    }

    private static boolean compareAmount(Row rowInSheet1, Row rowIn2) {

        String s1 = getStringValue(rowInSheet1.getCell(1));
        String s2 = getStringValue(rowIn2.getCell(1));

        if (s1 == null || s1.isBlank())
            return false;

        return s1.equals(s2);
    }

    private static boolean compareEntityLocation(Row rowInSheet1, Row rowIn2) {
        String s1 = getStringValue(rowInSheet1.getCell(2));
        String s2 = getStringValue(rowIn2.getCell(2));

        if (s1 == null || s1.isBlank())
            return false;

        return s1.equals(s2);
    }

    private static boolean compareTransit(Row rowInSheet1, Row rowIn2) {
        /*Cell c1 = rowInSheet1.getCell(22);
        Cell c2 = rowInSheet1.getCell(23);
        String s1= getStringValue(c1) + getStringValue(c2);*/
        String s1 = getStringValue(rowInSheet1.getCell(0));
        String s2 = getStringValue(rowIn2.getCell(0));

        if (s1 == null || s1.isBlank())
            return false;

        return s1.equals(s2);
        //return false;
    }

    private static boolean compareDatetime(Row rowInSheet1, Row rowIn2) {
/*
        String s1 = getStringValue(rowInSheet1.getCell(3));
        String s2 = getStringValue(rowIn2.getCell(3));

        if (s1 == null || s1.isBlank())
            return false;

        return s1.equals(s2);
*/
        return true;
    }

    private static String getStringValue(Cell c1) {
        String s = null;
        switch (c1.getCellType()) {
            case STRING:
                s = c1.getStringCellValue();
                break;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(c1)) {
                    s = c1.getDateCellValue().toString();
                } else {
                    s = Double.toString(c1.getNumericCellValue());
                }
                break;
            case BOOLEAN:
                //c1.getBooleanCellValue() ;
                System.out.println("Cell containing Boolean");
                break;
            case FORMULA:
                System.out.println("Cell containing FORMUAL");
                //c1.getCellFormula() ;
                break;
            default:
                System.out.println("Cell containing junk");
        }
        return s;
    }

    public static void writeExcel(XSSFWorkbook wb, FileOutputStream out) {

        try {
            wb.write(out);
            out.flush();
            out.close();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }

    }

    public static void getDatathroughExcel(String Sheetname, String excelonepath, String exceltwopath, String excelthreepath) throws IOException {
        //file input stream object
        FileInputStream inputStream1 = new FileInputStream(excelonepath);
        FileInputStream inputStream2 = new FileInputStream(exceltwopath);
        FileOutputStream outputStream = new FileOutputStream(excelthreepath);

        XSSFWorkbook workbook1 = new XSSFWorkbook(inputStream1);
        XSSFWorkbook workbook2 = new XSSFWorkbook(inputStream2);
        XSSFWorkbook workbook3 = new XSSFWorkbook();

        XSSFSheet x1 = workbook1.getSheet(Sheetname);
        XSSFSheet x2 = workbook2.getSheet(Sheetname);
        XSSFSheet x3 = workbook3.createSheet("report");

        int rowcount1 = x1.getPhysicalNumberOfRows();
        int rowcount2 = x2.getPhysicalNumberOfRows();

        System.out.println("Row counts File1:" + rowcount1 + ",File2:" + rowcount2);

        //FormulaEvaluator evaluator = workbook1.getCreationHelper().createFormulaEvaluator();
        boolean found = false;

        int rowcount = 0;

        XSSFRow row = x3.createRow(rowcount++);
        int col = 0;

        row.createCell(col++).setCellValue("DT_OPER");
        row.createCell(col++).setCellValue("HRE_OPER");
        row.createCell(col++).setCellValue("TRANSIT");
        row.createCell(col++).setCellValue("SUCCURSALE");
        row.createCell(col++).setCellValue("Branch");
        row.createCell(col++).setCellValue("Account");
        row.createCell(col).setCellValue("SUFFIXE");
        printRow(x2);
        for (int j = 1; j < rowcount1; j++) {
            Row rowInSheet1 = x1.getRow(j);

            found = searchInTie(rowInSheet1, x2);
            if (found) {
                System.out.println("found :" + j);
//                x1.removeRow(rowInSheet1);
            } else {
                XSSFRow diffrow = x3.createRow(rowcount++);
                int diffcol = 0;
                for (Cell cell : rowInSheet1) {
                    diffrow.createCell(diffcol++).setCellValue(getStringValue(cell));
                }
            }

        }

        writeExcel(workbook3, outputStream);

        if (workbook1 != null) workbook1.close();
        if (workbook2 != null) workbook2.close();

    }

    public static void printRow(XSSFSheet x1) {
        int rowcount1 = x1.getPhysicalNumberOfRows();
        for (int j = 1; j < rowcount1; j++) {
            Row rowInSheet1 = x1.getRow(j);

            System.out.print("Row:" + j);
            int c = 0;
            for (Cell cell : rowInSheet1) {
                //System.out.print("," + getStringValue(cell));
                System.out.print("," + c + "," + getStringValue(rowInSheet1.getCell(c++)));
            }
            System.out.println();
        }
    }
}
	
    
