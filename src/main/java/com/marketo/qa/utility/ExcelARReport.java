package com.marketo.qa.utility;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;





public class ExcelARReport {
	
	static String ExcelPath = System.getProperty("user.dir") + "//Reports//";
	static String fileName = new SimpleDateFormat("dd_MM_yy_HH_mm").format(new Date());

	public static void ARReport() throws Exception
	{
		
		
	String report = System.getProperty("user.dir") + "//Reports//"+passData.Exceldata("Account Name")+"_"+fileName +".xlsx";
	//FileInputStream FileInputStream = new FileInputStream(report);
	 
	            // Create a blank sheet
			 try {
		            // create blank workbook
		            XSSFWorkbook workbook = new XSSFWorkbook();
		            // Create a blank sheet
		            XSSFSheet AR_Data_Point1 = workbook.createSheet("AR Data Point 1");
		            XSSFSheet AR_Data_Point2 = workbook.createSheet("AR Data Point 2");
		            XSSFSheet CCT= workbook.createSheet("CCT ONLY");
			 
			 FileOutputStream out = new FileOutputStream(new File(report));
	            workbook.write(out);
	            out.close();
	            
	         // Close workbook
	            workbook.close();
			 }
			 catch (Exception e) {
				// TODO: handle exception
			}
	
	//	ExcelStyling.WriteExcelData( "test");
			
	}	
}
	 

