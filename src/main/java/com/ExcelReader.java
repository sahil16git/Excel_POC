package com;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReader {
	
	@SuppressWarnings("deprecation")
	public static String getCellValue(Cell cell, FormulaEvaluator evaluator){
		String returnCellValue="";
		CellValue cellValue = evaluator.evaluate(cell);

		switch (cellValue.getCellType()) {
		    case Cell.CELL_TYPE_BOOLEAN:
		    	returnCellValue=String.valueOf(cellValue.getBooleanValue());
		        break;
		    case Cell.CELL_TYPE_NUMERIC:
		    	returnCellValue=String.valueOf(cellValue.getNumberValue());
		        break;
		    case Cell.CELL_TYPE_STRING:
		    	returnCellValue=cellValue.getStringValue();
		        break;
		    case Cell.CELL_TYPE_BLANK:
		        break;
		    case Cell.CELL_TYPE_ERROR:
		        break;

		    // CELL_TYPE_FORMULA will never happen
		    case Cell.CELL_TYPE_FORMULA: 
		        break;
		}	
		return returnCellValue;
	}
	
	
	/**
	 * @param args
	 * @throws IOException 
	 */
	public static void main(String[] args) throws IOException {
		
        try {
        	System.out.println("Hi! Will read generic_excel_file.xlsx from resources");
        	Class cls = Class.forName("com.ExcelReader");
        	ClassLoader cLoader = cls.getClassLoader();
			InputStream stream = cLoader.getResourceAsStream("generic_excel_file.xlsx");
			XSSFWorkbook workbook = new XSSFWorkbook(stream);
			FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
			// Get first/desired sheet from the workbook
			XSSFSheet sheet = workbook.getSheetAt(0);
			Iterator<Row> rowIterator = sheet.iterator();
			while (rowIterator.hasNext()) {
				
				Row row = rowIterator.next();
				System.out.println(row.getRowNum());
				// Skipping the Title Row  or first row containing the mappings 
				if (row.getRowNum() == 0 || row.getRowNum() == 1 )
					continue;
				
				// Properties keys loaded to Enumeration
				
				// New Object
				
				String configId = getCellValue(row.getCell(0), evaluator).toUpperCase();
				System.out.println(configId);
			}
			System.out.println("Hello");
			workbook.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

}
