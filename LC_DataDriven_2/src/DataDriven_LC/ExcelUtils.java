package DataDriven_LC;

import java.io.FileInputStream;

import java.io.FileNotFoundException;

import java.io.FileOutputStream;

import java.io.IOException;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;

import org.apache.poi.xssf.usermodel.XSSFRow;

import org.apache.poi.xssf.usermodel.XSSFSheet;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtils {
	
	private static XSSFSheet ExcelWSheet;

	private static XSSFWorkbook ExcelWBook;

	private static XSSFCell Cell;

	private static XSSFRow Row;

public static Object[][] getTableArray(String FilePath, String SheetName) throws Exception {   

   String[][] tabArray = null;

   try {

	   FileInputStream ExcelFile = new FileInputStream(FilePath);

	   // Access the required test data sheet

	   ExcelWBook = new XSSFWorkbook(ExcelFile);

	   ExcelWSheet = ExcelWBook.getSheet(SheetName);

	   int startRow = 0;

	   int startCol = 0;

	   int ci,cj;

	   int totalRows = ExcelWSheet.getLastRowNum();

	   UpdateArrivalPort();

	   int totalCols = 6;

	   tabArray=new String[totalRows+1][totalCols];

	   ci=0;

	   for (int i=startRow;i<=totalRows;i++, ci++) {           	   

		  cj=0;

		   for (int j=startCol;j<totalCols;j++, cj++){

			   tabArray[ci][cj]=readExcel(i,j);

			   //System.out.println(tabArray[ci][cj]);  

				}

			}
	   ExcelFile.close();
	   //Create an object of FileOutputStream class to create write data in excel file

	    FileOutputStream outputStream = new FileOutputStream(FilePath);

	    //write data in the excel file

	    ExcelWBook.write(outputStream);

	    //close output stream

	    outputStream.close();

		}

	catch (FileNotFoundException e){

		System.out.println("Could not read the Excel sheet");

		e.printStackTrace();

		}

	catch (IOException e){

		System.out.println("Could not read the Excel sheet");

		e.printStackTrace();

		}

	return(tabArray);

	}

	public static void UpdateArrivalPort() throws Exception{
		try{

			Cell = ExcelWSheet.getRow(1).getCell(3);
			if(readExcel(1,3).equals("Chennai")){
				Cell.setCellValue("Agra");
			}
					
			
			}catch (Exception e){

			System.out.println(e.getMessage());

			throw (e);

			}
		}


	public static String readExcel(int RowNum, int ColNum) throws Exception {

		try{

			Cell = ExcelWSheet.getRow(RowNum).getCell(ColNum);
			
			DataFormatter formatter = new DataFormatter();
			String CellData = formatter.formatCellValue(Cell);
			return CellData;
			
			}catch (Exception e){

			System.out.println(e.getMessage());

			throw (e);

			}
		
	}
}
