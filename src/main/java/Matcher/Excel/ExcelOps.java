package Matcher.Excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.ParseException;
import java.util.Iterator;

public class ExcelOps {
    //Sheet name should be same
    public static boolean searchInTie(Row rowInSheet1, XSSFSheet sheet2) throws IOException {

        Iterator<Row> rowInSheet2 = sheet2.rowIterator();
        while (rowInSheet2.hasNext()) {
            Row rowIn2 = rowInSheet2.next();
            if ( compareDatetime(rowInSheet1,rowIn2) && compareTransit(rowInSheet1,rowIn2) &&  compareEntityLocation(rowInSheet1,rowIn2) &&  compareAmount(rowInSheet1,rowIn2))
             return true;        
        }
        
        return false;
        

    }

    private static boolean compareAmount(Row rowInSheet1, Row rowIn2) {

        return false;
    }

    private static boolean compareEntityLocation(Row rowInSheet1, Row rowIn2) {
        return false;
    }

    private static boolean compareTransit(Row rowInSheet1, Row rowIn2) {
        Cell c1 = rowInSheet1.getCell(22);
        Cell c2 = rowInSheet1.getCell(23);
        String s1= getStringValue(c1) + getStringValue(c2);
        String s2= getStringValue(rowIn2.getCell(4));
        return s1.equals(s2);
        //return false;
    }

    private static boolean compareDatetime(Row rowInSheet1, Row rowIn2) {
        return false;
    }
    private static String getStringValue(Cell c1)
    {
        String s = null;
        switch (c1.getCellType()) {
            case STRING:
                s = c1.getStringCellValue();
                break;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(c1)) {
                    s = c1.getDateCellValue().toString();
                } else {
                    s= Double.toString(c1.getNumericCellValue());
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
    public static void writeExcel( XSSFWorkbook wb ,String fileName) {

        try {
            File currDir = new File(".");
            String path = currDir.getAbsolutePath();
            String fileLocation = path.substring(0, path.length() - 1) + fileName + ".xlsx";
            FileOutputStream out = new FileOutputStream(fileLocation);


            wb.write(out);
            out.flush();
            out.close();
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }

    }

    public static void getDatathroughExcel(String Sheetname, String excelonepath, String exceltwopath, String excelthreepath) throws IOException, ParseException {
        //file input stream object
        FileInputStream inputStream1 = new FileInputStream(excelonepath);
        FileInputStream inputStream2 = new FileInputStream(exceltwopath);
       // FileInputStream inputStream3 = new FileInputStream(excelthreepath);

        XSSFWorkbook workbook1 = new XSSFWorkbook(inputStream1);
        XSSFWorkbook workbook2 = new XSSFWorkbook(inputStream2);
        //XSSFWorkbook workbook3 = new XSSFWorkbook(inputStream3);

        XSSFSheet x1 = workbook1.getSheet(Sheetname);
        XSSFSheet x2 = workbook2.getSheet(Sheetname);
        //XSSFSheet x3 = workbook3.getSheet(Sheetname);

        int rowcount1 = x1.getPhysicalNumberOfRows();
        int rowcount2 = x2.getPhysicalNumberOfRows();
     //   int rowcount3 = x3.getPhysicalNumberOfRows();

        System.out.println("Row counts: " + rowcount1 + "," + rowcount2 );

        //FormulaEvaluator evaluator = workbook1.getCreationHelper().createFormulaEvaluator();
        boolean found = false;
        for (int j = 1; j < rowcount1; j++) {

            Row rowInSheet1 = x1.getRow(j);
            found = searchInTie(rowInSheet1, x2);
            if (found) {
                System.out.println("Row :" + j);
                x1.removeRow(rowInSheet1);
            }
        }

        writeExcel(workbook1,excelthreepath);
        if (workbook1 != null) {
            workbook1.close();
        }
        if (workbook2 != null) {
            workbook2.close();
        }
        System.out.println("Hurray! work books diff completed....");
    }

}
	
    
