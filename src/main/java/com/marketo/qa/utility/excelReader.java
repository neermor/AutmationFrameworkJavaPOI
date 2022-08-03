package com.marketo.qa.utility;



import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class excelReader {
	
	
	 public static Map<String, String> getMapData() throws IOException 
	 {
		 HashMap<String, String>testData =new HashMap<String, String>();
		
		try {
			
	        String ExcelPath=System.getProperty("user.dir")+"./src/test/resources/TestData/MarketoData.xlsx";
			FileInputStream FileInputStream = new FileInputStream(ExcelPath);
			 try (Workbook workbook = new XSSFWorkbook(FileInputStream)) {
				Sheet sheet =workbook.getSheetAt(0);
				 int lastrownum = sheet.getLastRowNum();
				 
				 for(int i=1; i<lastrownum+1;i++)
				 {
					Row row =sheet.getRow(i); 
					Cell Keycell = row.getCell(0);
					
					String key = getCellValue(Keycell);	
					Cell Valuecell = row.getCell(1);
					String value = getCellValue(Valuecell);
					String valueformat=value.replaceAll("\\.0*$", "");      //Removing decimal value 
					testData.put(key, valueformat);
				 }
			}
		 }
		catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		 return testData;
		
	 }
	 
	 public static String getCellValue(Cell cell) {
		 String CellData = null;
		 
		 switch (cell.getCellType()) {
	        case  STRING:    //field that represents string cell type
	        	CellData= String.valueOf(cell.getStringCellValue()) ;
	        	break;
	        case NUMERIC:    //field that represents number cell type
	        	CellData=String.valueOf(cell.getNumericCellValue()) ;
	        	break;
	        case BOOLEAN:
	        	CellData= String.valueOf(cell.getBooleanCellValue());
	        	break;
	        default:
	            return "";
	    }
		 return CellData;
		}
}
	 
	
	




