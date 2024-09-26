package Matcher.Excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.text.ParseException;
import java.util.Date;

public class ExcelOps {
    //Sheet name should be same
    public static boolean isExistsTie(String Sheetname, String excelonepath, String exceltwopath) throws IOException {


        return true;

    }

    public static void getDatathroughExcel(String Sheetname, String excelonepath, String exceltwopath) throws IOException, ParseException {
        //file input stream object
        FileInputStream inputStream1 = new FileInputStream(excelonepath);
        FileInputStream inputStream2 = new FileInputStream(exceltwopath);
        XSSFWorkbook workbook1 = new XSSFWorkbook(inputStream1);
        XSSFWorkbook workbook2 = new XSSFWorkbook(inputStream2);
        XSSFSheet x1 = workbook1.getSheet(Sheetname);
        XSSFSheet x2 = workbook2.getSheet(Sheetname);
        int rowcount1 = x1.getPhysicalNumberOfRows();
        int rowcount2 = x2.getPhysicalNumberOfRows();
        System.out.println("row count:"+ rowcount1);
        System.out.println("row count:"+ rowcount2);

	/*	Assert.assertEquals(rowcount1,rowcount2, "Sheets have different count of rows..");
		Iterator<Row> rowInSheet1 = x1.rowIterator();
		Iterator<Row> rowInSheet2 = x2.rowIterator();
		while (rowInSheet1.hasNext()) {
			int cellCounts1 = rowInSheet1.next().getPhysicalNumberOfCells();
			int cellCounts2 = rowInSheet2.next().getPhysicalNumberOfCells();
			Assert.assertEquals(cellCounts1, cellCounts2, "Sheets have different count of columns..");
		}
	 */
        FormulaEvaluator evaluator = workbook1.getCreationHelper().createFormulaEvaluator();
        for (int j = 1; j < rowcount1; j++) {

            // Iterating through each cell
            int cellCounts = x1.getRow(j).getPhysicalNumberOfCells();
            //for (int k = 0; k < cellCounts; k++) {
            // Getting individual cell
			System.out.println("row *****************:"+ j+1);
            Cell c1 = x1.getRow(j).getCell(0);
            Cell c2 = x1.getRow(j).getCell(1);

            if (c1.getCellType() == CellType.NUMERIC) {
                // If cell type is numeric, we need to check if data is of Date type
                if (DateUtil.isCellDateFormatted(c1)) {
                    // Need to use DataFormatter to get data in given style otherwise it will come as time stamp
                    DataFormatter df = new DataFormatter();
                    //df.addFormat("dd/MM/yyyy", new java.text.SimpleDateFormat("yyyy/MM/dd"));
                    java.text.SimpleDateFormat sm = new java.text.SimpleDateFormat("yyyy/MM/dd");
                    java.text.SimpleDateFormat ind = new java.text.SimpleDateFormat("dd/MM/yyyy");
                    String v1 = df.formatCellValue(c1);
                    Date dt = sm.parse(v1);

                    //String v2 = df.formatCellValue(c2);
                    System.out.print("Date:" + v1);
                    System.out.println(","+ind.format(dt));
                } else {
						/*	double v1 = c1.getNumericCellValue();
							double v2 = c2.getNumericCellValue();
							Assert.assertEquals(v1, v2, "Cell values are different.....");*/
                    System.out.println("Its a number");
                }
            }
            if (c2.getCellType() == CellType.STRING) {
                String v2 = c2.getStringCellValue();
                v2 = v2.replaceAll(":","");
                System.out.println("TIME:" + v2);

            }

        }
        System.out.println("Hurray! work books diff completed....");
    }
}
	
    
