package com.marketo.qa.Pages;

import java.util.List;

import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;

import com.marketo.qa.FileLib.CommonLib;
import com.marketo.qa.base.TestBase;

public class DesignStudioPage extends TestBase{
	MyMarketoPage homepage= new MyMarketoPage();
	MarketingActivitePage mAP= new MarketingActivitePage();

	By AllCount = By.cssSelector("[data-id='dataGridFooter_pageInfo']");
	By UploadCount = By.cssSelector(".x-toolbar-right-row  .x-toolbar-cell .xtb-text");
	By Iframe = By.cssSelector("#mlm");
	By WorkSpace = By.xpath("//div[contains(@data-id,'treeNode_workspace')]/..//div//span");
	public WebElement GetIFrame() {
		return driver.findElement(Iframe);
	}
	
	public WebElement GetExpandBtn(String Name) {
		return driver.findElement(By.xpath("//div[contains(@data-id,'treeNode_workspace')]/..//div//span[text()='"+Name+"']/../preceding-sibling::button[contains(@data-id, 'treeNodeChevronIconButton')]//*[contains(@data-id, 'open')]"));
	}

	public int GetCount() throws Throwable {
		Thread.sleep(2000);
	    String countString = driver.findElement(AllCount).getText();
	    String[] words=countString.split("\\s"); 
	    if(words[0].equalsIgnoreCase("0")) {
	    	return 0;
	    }
	    return Integer.parseInt(words[4]);
	}
	
	
	
	public int GetUploadDataCount() throws Throwable {
		Thread.sleep(4000);
	    String countString = driver.findElement(UploadCount).getText();
	    String[] words=countString.split("\\s"); 	    
	    if(countString.equalsIgnoreCase("No images or files exist")) {
	    	return 0;
	    }
	    else {
		    return Integer.parseInt(words[1]);
		}
	}
	
	public int GetSnippetsCount() throws Throwable {
		Thread.sleep(10000);
	    String countString = driver.findElement(UploadCount).getText();
	    String[] words=countString.split("\\s");  
	    if(countString.equalsIgnoreCase("No snippets exist")) {
	    	return 0;
	    }
	    else {
		    return Integer.parseInt(words[0]);
		}
	}
	
	public int FetchTreeNodeCount(String value,int row,int cell) throws Throwable {
		homepage.SelectTreeNode(value);
		Thread.sleep(2000);
		new CommonLib().WriteExcelData("Sheet1", row, 0, value);
		new CommonLib().WriteExcelData("Sheet1", row, cell, GetCount());
		return GetCount();

	}
	
	public int FetchUploadCount(String value,int row,int cell) throws Throwable {
		driver.switchTo().defaultContent();
		homepage.SelectTreeNode(value);
		driver.switchTo().frame(GetIFrame());
		Thread.sleep(4000);
		new CommonLib().WriteExcelData("Sheet1", row, 0, value);
		new CommonLib().WriteExcelData("Sheet1", row, cell, GetUploadDataCount());	
		return GetUploadDataCount();

	}
	
	public int FetchSnippetsCount(String value,int row, int cell) throws Throwable {
		driver.switchTo().defaultContent();
		homepage.SelectTreeNode(value);
		driver.switchTo().frame(GetIFrame());
		Thread.sleep(4000);
		new CommonLib().WriteExcelData("Sheet1", row, 0, value);
		new CommonLib().WriteExcelData("Sheet1", row, cell, GetSnippetsCount());
		return GetSnippetsCount();

	}
	
	int cell = 2;

	int AllEmails = 0;
	int AllForms = 0;
	int AllLandingPages = 0;
	int AllImages_and_Files = 0;
	int AllSnippets = 0;

	
public void AllWorkspaceRequiredCount() throws Throwable {
	
		
		List<WebElement> workSpace = driver.findElements(WorkSpace);
		
		  workSpace.size();
		  new CommonLib().ClearExcelData("Sheet1", 1); 
		  new CommonLib().ClearExcelData("Sheet1", 2); 
		  new CommonLib().ClearExcelData("Sheet1", 3); 
		  new CommonLib().ClearExcelData("Sheet1", 4); 
		  new CommonLib().ClearExcelData("Sheet1", 5); 
		  new CommonLib().ClearExcelData("Sheet1", 6);
		 
		for(WebElement value : workSpace) {
			try {
				if(GetExpandBtn(value.getText()).isDisplayed()) {
					}
			}
			catch (Exception e) {
				Actions act = new Actions(driver);
				act.doubleClick(value).perform();					
			
			}
			
			  AllEmails += FetchTreeNodeCount("Emails",2,cell);			  
			  AllForms +=FetchTreeNodeCount("Forms", 3,cell);			  
			  AllLandingPages +=FetchTreeNodeCount("Landing Pages", 4,cell);			  
			  AllImages_and_Files +=FetchUploadCount("Images and Files", 5,cell);			  
			  AllSnippets +=FetchSnippetsCount("Snippets",6,cell);			 
			
			driver.switchTo().defaultContent();
			
			Actions act = new Actions(driver);
			act.doubleClick(value).perform();
	
			
			new CommonLib().WriteExcelData("Sheet1", 1, 0, "Asset Data");
			new CommonLib().WriteExcelData("Sheet1", 1, cell, value.getText());			
			cell++;		


		}

			  new CommonLib().WriteExcelData("Sheet1", 1, 1, "Total"); 
			  new CommonLib().WriteExcelData("Sheet1", 2, 1, AllEmails); 
			  new CommonLib().WriteExcelData("Sheet1", 3, 1, AllForms);
			  new CommonLib().WriteExcelData("Sheet1", 4, 1, AllLandingPages); 
			  new CommonLib().WriteExcelData("Sheet1", 5, 1, AllImages_and_Files); 
			  new CommonLib().WriteExcelData("Sheet1", 6, 1, AllSnippets); 
			 		


	}

	
	
	
	
	

}
