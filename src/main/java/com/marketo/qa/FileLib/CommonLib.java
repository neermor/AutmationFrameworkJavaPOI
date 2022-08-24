package com.marketo.qa.FileLib;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.time.Duration;
import java.util.concurrent.TimeUnit;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.FluentWait;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.Wait;
import org.openqa.selenium.support.ui.WebDriverWait;

import com.google.common.base.Function;
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
	 XSSFRow roww = null;

	public void WriteExcelData(String sheetName ,int row,int col,int cellValue) throws Exception {
        String ExcelPath=System.getProperty("user.dir")+"./src/test/resources/TestData/MarketoData.xlsx";
		File file =  new File(ExcelPath);
        FileInputStream fis = new FileInputStream(file); 
        XSSFWorkbook book = new XSSFWorkbook(fis);
        XSSFSheet sheet = book.getSheet(sheetName);
        
        roww = sheet.getRow(row);
        
        if(roww == null) {
        	sheet.createRow(row).createCell(col).setCellValue(cellValue);
        }
        else {
        	sheet.getRow(row).createCell(col).setCellValue(cellValue);

        }        
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
        roww = sheet.getRow(row);

        if(roww == null) {
        	sheet.createRow(row).createCell(col).setCellValue(cellValue);
        }
        else {
        	sheet.getRow(row).createCell(col).setCellValue(cellValue);

        }        FileOutputStream fos = new FileOutputStream(file);
        wb.write(fos);
        fos.close();
        
		}
	
	public void ClearExcelData(String sheetName ,int row) throws Exception {
		String ExcelPath=System.getProperty("user.dir")+"./src/test/resources/TestData/MarketoData.xlsx";
		File file =  new File(ExcelPath);
        FileInputStream fis = new FileInputStream(file); 
        XSSFWorkbook wb = new XSSFWorkbook(fis);
        XSSFSheet sheet = wb.getSheet(sheetName);
        int lastCell=sheet.getRow(row).getLastCellNum();
        for (int i = 0;i<lastCell;i++){
        	System.out.println(i);
        	sheet.createRow(row).createCell(i).setCellValue("");        	
        }
        FileOutputStream fos = new FileOutputStream(file);
        wb.write(fos);
        fos.close();

  	}
        
        
    
	
	
	public void MouseHover(WebElement element) {
		Actions actions = new Actions(driver);
		actions.moveToElement(element).perform();
	}
	
	public WebElement waitForElementVisibleFlunt( final By locator)
	{
		Wait<WebDriver> wait = new FluentWait<WebDriver>(driver)
				  .withTimeout(Duration.ofSeconds(20))
				  .pollingEvery(Duration.ofSeconds(3))
				  .withMessage("Waiting for element to load")
				  .ignoring(NoSuchElementException.class);
		
		WebElement element = wait.until(new Function<WebDriver, WebElement>()
		  {
		    public WebElement apply(WebDriver driver)
		    {
		      return driver.findElement(locator);
		    }
		  });
		return element;
							
	}


	

       

	}
	
	
