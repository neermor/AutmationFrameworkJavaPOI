package com.marketo.qa.Pages;

import java.util.List;

import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;

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
	
	public WebElement GetIFrame() {
		return driver.findElement(Iframe);
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
	
	public String GetCount() throws Throwable {
		
		Thread.sleep(4000);
	    String countString = driver.findElement(LeadsCount).getText();
	    String[] words=countString.split("\\s"); 
	    if(words[0].equalsIgnoreCase("No")) {
			return "0";
		}
	    else {
	    return words[2];
	    }	
    
}
	
	public void SelectTreeNode(String TreeNodeName) throws Throwable {
		Thread.sleep(10000);
	    List<WebElement> HomeTiles = driver.findElements(TreeNode);
		 for(WebElement option: HomeTiles){
			 
				if (option.getText().equalsIgnoreCase(TreeNodeName)) {
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
	
	public void switchFrame() throws Throwable {
	}
	
	
	public void GetLeadsCount(int row) throws Throwable {
		new CommonLib().StandardWait(2000);
		driver.switchTo().frame(GetIFrame());
		GetPeople().click();
		new CommonLib().WriteExcelData("Sheet1", row, 0, "Leads");
		new CommonLib().WriteExcelData("Sheet1", row, 1, GetCount());
	}
	
	public void SegmentationsCount(int row) throws Throwable {
		try {
			boolean  flag=driver.findElement(By.xpath("//div[contains(@data-id,'treeNode_Label')]/span[text()='Segmentations']/../preceding-sibling::button[@data-id='treeNodeChevronIconButton']")).isDisplayed();
			System.out.println(flag);
			if(flag) {
			homePage.ExtendTreeNode("Segmentations");
			new CommonLib().WriteExcelData("Sheet1", row, 0, "Segmentations");
			new CommonLib().WriteExcelData("Sheet1", row, 1, GetSag().size());
			new CommonLib().StandardWait(2000);
			screenshotUtility.TakeScreenshot(GetSagHar(), "Segmentations");		 		
			new CommonLib().StandardWait(2000);
		}
			}
			catch (Exception e) {
				new CommonLib().WriteExcelData("Sheet1", row, 0, "Segmentations");
				new CommonLib().WriteExcelData("Sheet1", row, 1, 0);
				}
	}

}
