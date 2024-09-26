package Matcher.Excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.Assert;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

public class ExcelOpsOrg {
	//Sheet name should be same
	public static void getDatathroughExcel(String Sheetname,String excelonepath,String exceltwopath) throws IOException 
	{
		//file input stream object
		FileInputStream inputStream1 = new FileInputStream(excelonepath);
		FileInputStream inputStream2 = new FileInputStream(exceltwopath);
		XSSFWorkbook workbook1=new XSSFWorkbook(inputStream1);
		XSSFWorkbook workbook2=new XSSFWorkbook(inputStream2);
		XSSFSheet x1=workbook1.getSheet(Sheetname);
		XSSFSheet x2=workbook2.getSheet(Sheetname);
		int rowcount1=x1.getPhysicalNumberOfRows();
		int rowcount2=x2.getPhysicalNumberOfRows();
		Assert.assertEquals(rowcount1,rowcount2, "Sheets have different count of rows..");
		Iterator<Row> rowInSheet1 = x1.rowIterator();
		Iterator<Row> rowInSheet2 = x2.rowIterator();
		while (rowInSheet1.hasNext()) {
			int cellCounts1 = rowInSheet1.next().getPhysicalNumberOfCells();
			int cellCounts2 = rowInSheet2.next().getPhysicalNumberOfCells();
			Assert.assertEquals(cellCounts1, cellCounts2, "Sheets have different count of columns..");
		}
		for (int j = 0; j < rowcount1; j++) {
			// Iterating through each cell
			int cellCounts = x1.getRow(j).getPhysicalNumberOfCells();
			for (int k = 0; k < cellCounts; k++) {
				// Getting individual cell
				Cell c1 = x1.getRow(j).getCell(k);
				Cell c2 = x2.getRow(j).getCell(k);
				// Since cell have types and need o use different methods
				if (c1.getCellType().equals(c2.getCellType())) {
					if (c1.getCellType() == CellType.STRING) {
						String v1 = c1.getStringCellValue();
						String v2 = c2.getStringCellValue();
						Assert.assertEquals(v1, v2, "Cell values are different.....");
						System.out.println("Its matched : "+ v1 + " === "+ v2);
					}
					if (c1.getCellType() == CellType.NUMERIC) {
						// If cell type is numeric, we need to check if data is of Date type
						if (DateUtil.isCellDateFormatted(c1) | DateUtil.isCellDateFormatted(c2)) {
							// Need to use DataFormatter to get data in given style otherwise it will come as time stamp
							DataFormatter df = new DataFormatter();
							String v1 = df.formatCellValue(c1);
							String v2 = df.formatCellValue(c2);
							Assert.assertEquals(v1, v2, "Cell values are different.....");
							System.out.println("Its matched : "+ v1 + " === "+ v2);
						} else {
							double v1 = c1.getNumericCellValue();
							double v2 = c2.getNumericCellValue();
							Assert.assertEquals(v1, v2, "Cell values are different.....");
							System.out.println("Its matched : "+ v1 + " === "+ v2);
						}
					}
					if (c1.getCellType() == CellType.BOOLEAN) {
						boolean v1 = c1.getBooleanCellValue();
						boolean v2 = c2.getBooleanCellValue();
						Assert.assertEquals(v1, v2, "Cell values are different.....");
						System.out.println("Its matched : "+ v1 + " === "+ v2);
					}
				} else
				{
					// If cell types are not same, exit comparison  
					Assert.fail("Non matching cell type.");
				}
				
			}
			
		}
		System.out.println("Hurray! Both work books have same data.");
	    }
	    }
	
    
