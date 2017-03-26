package com.gogate.apps.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

public class ExcelReader {

	/*public static void main(String[] args){
		File excelFile = new File("C:\\Users\\gogate\\Desktop\\ps.xls");
		//readExcel(excelFile);
	}
	*/
	public void readExcel(File excelFile){
		
		FileInputStream fip;
		HSSFWorkbook wb;
		HSSFSheet ws;
		HSSFRow row;
		try {		
			fip = new FileInputStream(excelFile);
			wb = new HSSFWorkbook(fip);
			ws = wb.getSheet("Sheet1");
			Iterator<Row> rowIterator = ws.iterator();
			  while (rowIterator.hasNext()) 
		      {
		         row = (HSSFRow) rowIterator.next();
		         Iterator<Cell> cellIterator = row.cellIterator();
		         while ( cellIterator.hasNext()) 
		         {
		            Cell cell = cellIterator.next();
		            switch (cell.getCellType()) 
		            {
		               case Cell.CELL_TYPE_NUMERIC:
			               System.out.print(cell.getNumericCellValue() + " \t\t " );
			               break;
		               case Cell.CELL_TYPE_STRING:
			               System.out.print(cell.getStringCellValue() + " \t\t " );
			               break;
			        }
		         }
		         System.out.println();
		      }
		} catch (Exception e) {
			System.out.println(excelFile.getAbsolutePath());
			e.printStackTrace();
		}
		
	}
	
	public static void readExcelSpecific(File excelFile){
		
		FileInputStream fip;
		HSSFWorkbook wb;
		HSSFSheet ws;
		HSSFRow row;
		try {		
			fip = new FileInputStream(excelFile);
			wb = new HSSFWorkbook(fip);
			ws = wb.getSheet("Sheet1");
			Iterator<Row> rowIterator = ws.iterator();
			  while (rowIterator.hasNext()) 
		      {
		         row = (HSSFRow) rowIterator.next();
		         Iterator<Cell> cellIterator = row.cellIterator();
		         while ( cellIterator.hasNext()) 
		         {
		            Cell cell = cellIterator.next();
		            switch (cell.getCellType()) 
		            {
		               case Cell.CELL_TYPE_NUMERIC:
			               System.out.print(cell.getNumericCellValue() + " \t\t " );
			               break;
		               case Cell.CELL_TYPE_STRING:
			               System.out.print(cell.getStringCellValue() + " \t\t " );
			               break;
			        }
		         }
		         System.out.println();
		      }
		} catch (Exception e) {
			System.out.println(excelFile.getAbsolutePath());
			e.printStackTrace();
		}
		
	}

}
