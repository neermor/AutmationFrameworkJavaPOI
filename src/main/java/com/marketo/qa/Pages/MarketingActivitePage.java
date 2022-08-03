package com.marketo.qa.Pages;

import java.util.List;

import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;

import com.marketo.qa.FileLib.CommonLib;
import com.marketo.qa.base.TestBase;

public class MarketingActivitePage extends TestBase {
	
	boolean flag = false;
	By TreeNode=By.xpath("//div[contains(@data-id,'treeNodeRow' )]");
	By Iframe = By.cssSelector("#mlm");
	By CampaignInspector = By.cssSelector("#canvas__cp_campaignInspector");
	By CampaignDD= By.xpath("//button[text()='Active Campaigns']");
	By CampaignInspectorOption = By.cssSelector("[class='x-menu x-menu-floating x-layer'] ul li");
	By AllCount = By.cssSelector("[class='x-toolbar-right-ct'] [class='x-toolbar-cell'] div");
	By MoreCampaigns = By.xpath("//span[contains(text(),'More Campaigns')]");
	By MoreCampaignOptions = By.cssSelector("[class='x-menu x-menu-floating x-layer mktSubMenu'] ul li span");

	
	public WebElement GetIFrame() {
		return driver.findElement(Iframe);
	}
		
	public WebElement GetCampaignDD() {
		return driver.findElement(CampaignDD);
	}
	
	public WebElement GetCampaignInspector() {
		return driver.findElement(CampaignInspector);
	}
	
	public WebElement GetMoreCampaign() {
		return driver.findElement(MoreCampaigns);
	}
	
	public void switchFrame() throws Throwable {
		driver.switchTo().frame(GetIFrame());

	}
	
	
	public void ClickCampaignInspector() throws Throwable {
		new CommonLib().StandardWait(2000);
		GetCampaignInspector().click();	
		GetCampaignDD().click();
		
	} 
	
	public void SelectTreeNode(String TreeNodeName) throws Throwable {
		Thread.sleep(10000);
	    List<WebElement> HomeTiles = driver.findElements(TreeNode);
        System.out.println(HomeTiles);
		 for(WebElement option: HomeTiles){
			 
				if (option.getText().equalsIgnoreCase(TreeNodeName)) {
					option.click();
				}
				System.out.println(option.getText());
			 }
		 }
	

	public void GetMoreCampaignOption(String MoreCampaignValue) throws Throwable {
		Thread.sleep(10000);
	    List<WebElement> CampaignOption = driver.findElements(MoreCampaignOptions);
	    		System.out.println(CampaignOption);	    		
		 for(WebElement option: CampaignOption){			 
				if (option.getText().equalsIgnoreCase(MoreCampaignValue)) {
					option.click();
				}
				System.out.println(option.getText());
			 }	
		 }
	

	public void GetCampaignInspectorOption(String CampaignType) throws Throwable {
		Thread.sleep(10000);
	    List<WebElement> CampaignOption = driver.findElements(CampaignInspectorOption);
        System.out.println(CampaignOption);
		 for(WebElement option: CampaignOption){
			 
				if (option.getText().equalsIgnoreCase(CampaignType)) {
					option.click();
				}
				System.out.println(option.getText());
			 }	
		 }
	
	public void SelectWorkSpace(String WorkspaceName) throws Throwable {
		SelectTreeNode(WorkspaceName);
	}
	
	public String GetCount() throws Throwable {
		Thread.sleep(4000);
	    String countString = driver.findElement(AllCount).getText();
	    String[] words=countString.split("\\s");  
	    System.out.println(words[0]); 
	    return words[0];
	}
	
	public void GetCampaignCount(String CampaignType,int row)throws Throwable {
		ClickCampaignInspector();
		GetCampaignInspectorOption(CampaignType);
		new CommonLib().StandardWait(4000); 
		new CommonLib().WriteExcelData("Sheet1", row, 0, CampaignType);
		new CommonLib().WriteExcelData("Sheet1", row, 1, GetCount());
	}
	
	public void HoverMoreCampaign() throws Throwable{
		new CommonLib().MouseHover(GetMoreCampaign());
		new CommonLib().StandardWait(4000); 
	}
	
	public void GetMoreCampaignCount(String CampaignType,int row) throws Throwable {
		ClickCampaignInspector();
		HoverMoreCampaign();		
		GetMoreCampaignOption(CampaignType);
		new CommonLib().StandardWait(4000); 
		new CommonLib().WriteExcelData("Sheet1", row, 0, CampaignType);
		new CommonLib().WriteExcelData("Sheet1", row, 1, GetCount());


	}
	
	
	
	
	

}
