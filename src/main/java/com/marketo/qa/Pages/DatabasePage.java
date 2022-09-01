package com.marketo.qa.Pages;

import java.util.List;

import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;

import com.marketo.qa.FileLib.CommonLib;
import com.marketo.qa.base.TestBase;
import com.marketo.qa.utility.screenshotUtility;

public class DatabasePage extends TestBase{
	MyMarketoPage homePage= new MyMarketoPage();

	By TreeNode=By.xpath("//div[contains(@data-id,'treeNodeRow' )]");
	By Iframe = By.cssSelector("#mlm");
	By People = By.id("canvas__cp_ldbCanvasLeadList");
	By LeadsCount = By.cssSelector("[class='x-toolbar-right-row'] [class='xtb-text']");
	By Sag = By.xpath("//div[contains(@data-id,'treeNode_segmentation')]");
	By SagHar = By.cssSelector("#treeBodyAnchor > div > div > div:nth-child(6)"); 
	By WorkSpace = By.xpath("//div[contains(@data-id,'treeNode_workspace')]/..//div//span");

	public WebElement GetIFrame() {
		return driver.findElement(Iframe);
	}
	public WebElement GetExpandBtn(String Name) {
		return driver.findElement(By.xpath("//div[contains(@data-id,'treeNode_workspace')]/..//div//span[text()='"+Name+"']/../preceding-sibling::button[contains(@data-id, 'treeNodeChevronIconButton')]//*[contains(@data-id, 'open')]"));
	}

	
	public WebElement GetPeople() {
		return driver.findElement(People);
	}
	
	public List<WebElement> GetSag() {
		return driver.findElements(Sag);
	}
	
	public WebElement GetSagHar() {
		return driver.findElement(SagHar);
	}
	
	public int GetCount() throws Throwable {
		
		Thread.sleep(4000);
	    String countString = driver.findElement(LeadsCount).getText();
	    String[] words=countString.split("\\s"); 
	    if(words[0].equalsIgnoreCase("No")) {
			return 0;
		}
	    else {
	    return Integer.parseInt(words[2]);
	    }	
    
}
	
	public void SelectTreeNode(String TreeNodeName) throws Throwable {
		Thread.sleep(10000);
	    List<WebElement> HomeTiles = driver.findElements(TreeNode);
		 for(WebElement option: HomeTiles){
			 
				if (option.getText().startsWith(TreeNodeName)) {
					option.click();
				}
			 }
		 }
	

	
	public void  ExtendWorkshoptreenode(String Name,String ListName) throws Throwable {
		new CommonLib().StandardWait(2000);
		driver.findElement(By.xpath("//div[contains(@data-id,'treeNode_Label')]/span[text()="+"'"+Name+"'"+"]/../preceding-sibling::button[@data-id='treeNodeChevronIconButton']")).click();
		SelectTreeNode(ListName);
		new CommonLib().StandardWait(2000);
	}
	
	
	
	public int GetLeadsCount(int row,int cell) throws Throwable {
		new CommonLib().StandardWait(2000);
		driver.switchTo().frame(GetIFrame());
		GetPeople().click();
		new CommonLib().WriteExcelData("Sheet1", row, 0, "Leads");
		new CommonLib().WriteExcelData("Sheet1", row, cell, GetCount());
		return GetCount();
	}
	
	public int SegmentationsCount(int row,int cell) throws Throwable {
		try {
			boolean  flag=driver.findElement(By.xpath("//div[contains(@data-id,'treeNode_Label')]/span[text()='Segmentations']/../preceding-sibling::button[@data-id='treeNodeChevronIconButton']")).isDisplayed();
			System.out.println(flag);
			if(flag) {
			homePage.ExtendTreeNode("Segmentations");
			new CommonLib().WriteExcelData("Sheet1", row, 0, "Segmentations");
			new CommonLib().WriteExcelData("Sheet1", row, cell, GetSag().size());
			new CommonLib().StandardWait(2000);
			screenshotUtility.TakeScreenshot(GetSagHar(), "Segmentations"+cell);	
			return GetSag().size();
		}
			}
			catch (Exception e) {
				new CommonLib().WriteExcelData("Sheet1", row, 0, "Segmentations");
				new CommonLib().WriteExcelData("Sheet1", row, cell, 0);
				}
		return 0;

	}
	
	
	
	int cell = 2;

	int Segmentations = 0;
	int Leads = 0;
	
	public void AllWorkspaceCollectLeadsCount() throws Throwable {
		
		List<WebElement> workSpace = driver.findElements(WorkSpace);
		new CommonLib().ClearExcelData("Sheet1", 14);
		new CommonLib().ClearExcelData("Sheet1", 15);
		new CommonLib().ClearExcelData("Sheet1", 16);
		
		for(WebElement value : workSpace) {
			try {
				System.out.println(GetExpandBtn(value.getText()).isDisplayed());
				if(GetExpandBtn(value.getText()).isDisplayed()) {
					}
			}
			catch (Exception e) {
				Actions act = new Actions(driver);
				act.doubleClick(value).perform();					
			
			}
			Segmentations +=SegmentationsCount(15,cell);
			ExtendWorkshoptreenode("System Smart Lists","All People");
			Leads +=GetLeadsCount(16,cell);
			
			driver.switchTo().defaultContent();
			Actions act = new Actions(driver);
			act.doubleClick(value).perform();
	
			
			new CommonLib().WriteExcelData("Sheet1", 14, 0, "Database Data");
			new CommonLib().WriteExcelData("Sheet1", 14, cell, value.getText());			
			cell++;		


		} 
			  new CommonLib().WriteExcelData("Sheet1", 14, 1, "Total"); 
			  new CommonLib().WriteExcelData("Sheet1", 15, 1, Segmentations); 
			  new CommonLib().WriteExcelData("Sheet1", 16, 1, Leads);
			  
			  
			  


	}


}
