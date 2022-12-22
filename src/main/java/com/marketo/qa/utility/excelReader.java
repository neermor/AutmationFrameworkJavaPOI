package com.marketo.qa.utility;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
<<<<<<< HEAD
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
=======
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;
>>>>>>> 756f7ccb649fbee02a08964e9b289fc0f41ab3a3
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class excelReader {

<<<<<<< HEAD
	//fetching data from excel sheet
=======
>>>>>>> 756f7ccb649fbee02a08964e9b289fc0f41ab3a3
	public static Map<String, String> getMapData(String sheetName) throws IOException {
		HashMap<String, String> testData = new HashMap<String, String>();
		
		try {

			String ExcelPath = System.getProperty("user.dir") + "//Config//MarketoData.xlsx";
			File file = new File(ExcelPath);
			FileInputStream FileInputStream = new FileInputStream(ExcelPath);
			try (Workbook workbook = new XSSFWorkbook(FileInputStream)) {
				Sheet sheet = workbook.getSheet(sheetName);
				int lastrownum = sheet.getLastRowNum();
				
				for (int i = 0; i<lastrownum+1; i++) {
					Row row = sheet.getRow(i);

					if (row == null) {
						sheet.createRow(i).createCell(0).setCellValue(" ");
						FileOutputStream fos = new FileOutputStream(file);
						workbook.write(fos);
					}


					Cell Keycell = row.getCell(0, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
					String key = getCellValue(Keycell);

					Cell Valuecell = row.getCell(1, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
					String value = getCellValue(Valuecell);

					String valueformat = value.replaceAll("\\.0*$", ""); // Removing decimal value
					testData.put(key, valueformat);
				}
			}

		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return testData;

	}
<<<<<<< HEAD
	//method overloading for excel sheet
=======
	
>>>>>>> 756f7ccb649fbee02a08964e9b289fc0f41ab3a3
	public static Map<String, String> getMapData(int cell) throws IOException {
		HashMap<String, String> testData = new HashMap<String, String>();

		try {

			String ExcelPath = System.getProperty("user.dir") + "//Config//MarketoData.xlsx";
			FileInputStream FileInputStream = new FileInputStream(ExcelPath);
			try (Workbook workbook = new XSSFWorkbook(FileInputStream)) {
				Sheet sheet = workbook.getSheetAt(0);
				int lastrownum = sheet.getLastRowNum();
				
				for (int i = 0; i < lastrownum+1 ; i++) {
					Row row = sheet.getRow(i);

					Cell Keycell = row.getCell(0, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
					String key = getCellValue(Keycell);

					Cell Valuecell = row.getCell(cell,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
					
					String value =getCellValue(Valuecell);
					
					String valueformat = value.replaceAll("\\.0*$", ""); // Removing decimal value
					testData.put(key, valueformat);
				}
			}

		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return testData;

	}

<<<<<<< HEAD
	// adding switch case which perform fetching for different type of values 
=======
	
		
	

>>>>>>> 756f7ccb649fbee02a08964e9b289fc0f41ab3a3
	public static String getCellValue(Cell cell) {
		String CellData = null;

		switch (cell.getCellType()) {
		case STRING: // field that represents string cell type
			CellData = String.valueOf(cell.getStringCellValue());
			break;
		case NUMERIC: // field that represents number cell type
			CellData = String.valueOf(cell.getNumericCellValue());
			break;
		case BOOLEAN:
			CellData = String.valueOf(cell.getBooleanCellValue());
			break;
		case BLANK:
			CellData = String.valueOf(cell.getRichStringCellValue());
			break;

		default:
			return "";
		}
		return CellData;
	}

<<<<<<< HEAD
	//Different Approach for Excel model reader 
=======
>>>>>>> 756f7ccb649fbee02a08964e9b289fc0f41ab3a3
	public static String excelModelReader(int row, int cell) throws NumberFormatException, IOException {
		String ExcelPath = System.getProperty("user.dir") + "//Config//MarketoData.xlsx";
		FileInputStream FileInputStream = new FileInputStream(ExcelPath);
		try (Workbook workbook = new XSSFWorkbook(FileInputStream)) {
			Sheet sheet = workbook.getSheetAt(0);

			String data = sheet.getRow(row).getCell(cell).toString();
			return data;
		}
	}

	

}
