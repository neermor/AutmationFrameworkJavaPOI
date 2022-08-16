package com.marketo.qa.Pages;

import java.util.List;

import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;

import com.marketo.qa.FileLib.CommonLib;
import com.marketo.qa.base.TestBase;

public class DesignStudioPage extends TestBase{
	MyMarketoPage homepage= new MyMarketoPage();
	
	By AllCount = By.cssSelector("[data-id='dataGridFooter_pageInfo']");
	By UploadCount = By.cssSelector(".x-toolbar-right-row  .x-toolbar-cell .xtb-text");
	By Iframe = By.cssSelector("#mlm");

	public WebElement GetIFrame() {
		return driver.findElement(Iframe);
	}

	public String GetCount() throws Throwable {
		Thread.sleep(2000);
	    String countString = driver.findElement(AllCount).getText();
	    String[] words=countString.split("\\s");  
	    return words[4];
	}
	
	public String GetUploadDataCount() throws Throwable {
		Thread.sleep(4000);
	    String countString = driver.findElement(UploadCount).getText();
	    String[] words=countString.split("\\s"); 	    
	    if(countString.equalsIgnoreCase("No images or files exist")) {
	    	return "0";
	    }
	    else {
		    return words[1];
		}
	}
	
	public String GetSnippetsCount() throws Throwable {
		Thread.sleep(10000);
	    String countString = driver.findElement(UploadCount).getText();
	    String[] words=countString.split("\\s");  
	    if(countString.equalsIgnoreCase("No snippets exist")) {
	    	return "0";
	    }
	    else {
		    return words[0];
		}
	}
	
	public void FetchTreeNodeCount(String value,int row) throws Throwable {
		homepage.SelectTreeNode(value);
		Thread.sleep(2000);
		new CommonLib().WriteExcelData("Sheet1", row, 0, value);
		new CommonLib().WriteExcelData("Sheet1", row, 1, GetCount());	

	}
	
	public void FetchUploadCount(String value,int row) throws Throwable {
		driver.switchTo().defaultContent();
		homepage.SelectTreeNode(value);
		driver.switchTo().frame(GetIFrame());
		Thread.sleep(4000);
		new CommonLib().WriteExcelData("Sheet1", row, 0, value);
		System.out.println(GetUploadDataCount());
		new CommonLib().WriteExcelData("Sheet1", row, 1, GetUploadDataCount());	

	}
	
	public void FetchSnippetsCount(String value,int row) throws Throwable {
		driver.switchTo().defaultContent();
		homepage.SelectTreeNode(value);
		driver.switchTo().frame(GetIFrame());
		Thread.sleep(4000);
		new CommonLib().WriteExcelData("Sheet1", row, 0, value);
	    System.out.println(GetSnippetsCount()); 
		new CommonLib().WriteExcelData("Sheet1", row, 1, GetSnippetsCount());	

	}
	
	
	

}
