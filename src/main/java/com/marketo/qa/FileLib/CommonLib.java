package com.marketo.qa.FileLib;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.concurrent.TimeUnit;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;

import com.marketo.qa.base.TestBase;

public class CommonLib extends TestBase{
	 
	
	public  void waitForPageLoad(int time) {
		driver.manage().timeouts().implicitlyWait(time,TimeUnit.SECONDS);
	}
	
	public void WaitForElementToLoad(int duration,WebElement element) {
		WebDriverWait wait = new WebDriverWait(driver, duration);
		wait.until(ExpectedConditions.elementToBeClickable(element));
	}
	
	public void SelectDropDownValue(WebElement element,String selectValue ) {
		Select dropDown = new Select(element);
		dropDown.selectByVisibleText(selectValue);
	}
	
	public void StandardWait(int time) throws Throwable {
		Thread.sleep(time);
	}
	
	public void WriteExcelData(String sheetName ,int row,int col,int cellValue) throws Exception {
        String ExcelPath=System.getProperty("user.dir")+"./src/test/resources/TestData/MarketoData.xlsx";
		File file =  new File(ExcelPath);
        FileInputStream fis = new FileInputStream(file); 
        XSSFWorkbook book = new XSSFWorkbook(fis);
        XSSFSheet sheet = book.getSheet(sheetName);
        sheet.getRow(row).createCell(col).setCellValue(cellValue);
        FileOutputStream fos = new FileOutputStream(file);
        book.write(fos);
        fos.close();

	}
	
	public void WriteExcelData(String sheetName ,int row,int col,String cellValue) throws Exception {
        String ExcelPath=System.getProperty("user.dir")+"./src/test/resources/TestData/MarketoData.xlsx";
		File file =  new File(ExcelPath);
        FileInputStream fis = new FileInputStream(file); 
        XSSFWorkbook wb = new XSSFWorkbook(fis);
        XSSFSheet sheet = wb.getSheet(sheetName);
        sheet.getRow(row).createCell(col).setCellValue(cellValue);
        FileOutputStream fos = new FileOutputStream(file);
        wb.write(fos);
        fos.close();
        
		}
	
	public void MouseHover(WebElement element) {
		Actions actions = new Actions(driver);
		actions.moveToElement(element).perform();
	}
	

       

	}
	
	
