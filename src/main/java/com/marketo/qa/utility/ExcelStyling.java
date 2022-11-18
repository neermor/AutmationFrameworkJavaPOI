package com.marketo.qa.utility;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelStyling {
	
	
	static String ExcelPath = System.getProperty("user.dir") + "//Reports//";
	static String fileName = new SimpleDateFormat("dd_MM_yy_HH_mm").format(new Date());
	
	
	
	public static void WriteExcelData(XSSFWorkbook workbook) throws Exception
	{
		
			
			
		String report = System.getProperty("user.dir") + "//Reports//"+passData.Exceldata("Account Name")+"_"+fileName +".xlsx";
		
		 
		        
						
						
			            // create blank workbook
							workbook = new XSSFWorkbook(report);
							XSSFSheet AR_Data_Point4 = workbook.getSheet("test");
			            // Create a blank sheet
			            XSSFSheet AR_Data_Point1 = workbook.createSheet("AR_Data_Point1");
			            XSSFSheet AR_Data_Point2 = workbook.createSheet("AR_Data_Point2");
			            XSSFSheet CCT= workbook.createSheet("CCT ONLY");
				 
			           
						 FileOutputStream out = new FileOutputStream(new File(report));
				            workbook.write(out);
				            out.close();
		            
		         // Close workbook
		            workbook.close();
				 }
				 
		//	ExcelStyling.WriteExcelData( "test");
				


}
