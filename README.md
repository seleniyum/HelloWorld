# HelloWorld
--------------------------------------------EXcel Reader -----------------------------------------------
package excelWorks;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReadWrite {

	public FileInputStream fis = null; //Declaring a File Input Stream
	//private XSSFWorkbook workbook = null; // Declaring workbook to work with
//	private XSSFSheet sheet = null ; 	// Declaring a Sheet to work with
	//private XSSFRow row = null; 
	private HSSFCell cell = null;
	private  HSSFWorkbook workbook = null;
	private HSSFSheet sheet = null ; 	// Declaring a Sheet to work with
	private HSSFRow row = null; 
	String path = null; //- This could be used to be used with fis -> fis = new FileInputStream(path);
	//if  equated to the path of excel sheet. 
	
	
	public ExcelReadWrite() throws  IOException {
		
		path = System.getProperty("user.dir")+"\\exlib\\TestSheet.xls";
		fis = new FileInputStream(path); //feeding path to fis
		workbook = new HSSFWorkbook(fis); 		// passing fis to workbook
		sheet = workbook.getSheetAt(0);			// taking the first sheet at index 0
		
	}
	
	public int getSheetRows(String sheetName) { // creating a class to work with Sheet Row count
	
		int index = workbook.getSheetIndex(sheetName);// creating index to use sheetName to find index.
		sheet = workbook.getSheetAt(index);		// performing a sheet selection from the index. 
		
		return (sheet.getLastRowNum()+1);		// performing return value from sheet's total row count
	
	}
	public int getSheetCol(String sheetName) { // creating a class to work with Sheet Column count
		
		int index = workbook.getSheetIndex(sheetName);// creating index to use sheetName to find index.
		sheet = workbook.getSheetAt(index);		// performing a sheet selection from the index. 
		row = sheet.getRow(0);
		
		return (sheet.getLastRowNum() );		// performing return value from sheet's total row count
	
	}
	
	public String getCellData(String sheetName, int colNum, int rowNum){
		int index = workbook.getSheetIndex(sheetName);
		sheet = workbook.getSheetAt(index);
		row = sheet.getRow(rowNum);
		cell = row.getCell(colNum);
		
		return (cell.getStringCellValue());
		
			}
	
	
	public static void main(String[] args) throws IOException {
		
		ExcelReadWrite reader = new ExcelReadWrite();
		System.out.println(reader.getSheetRows("SearchApt"));
		System.out.println(reader.getSheetCol("SearchApt"));
		
		System.out.println(reader.getCellData("SearchApt", 1, 1)); 
		

	}

}
