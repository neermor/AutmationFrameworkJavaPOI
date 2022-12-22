package com.marketo.qa.Pages;

import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;

import com.marketo.qa.FileLib.CommonLib;
import com.marketo.qa.base.TestBase;

public class SupportToolsPage extends TestBase{
	
	private By flowAction=By.cssSelector("#actionId");
	private By listCampaigns = By.xpath("//button[text()='List Campaigns']");
	private By actionCount = By.xpath("//h3[contains(text(),'Campaign Flows with')]/..//table/tbody/tr");
	private By resultTable = By.cssSelector("[name='results'] table");
	private By backToSupport = By.xpath("//a[contains(text(),'back to Support Tools')]");
	public WebElement GetFlowAction() {
		return driver.findElement(flowAction);
	}
	
	public WebElement GetListCampaign() {
		return driver.findElement(listCampaigns);
	}
	
	public WebElement GetBackToSupportLink() {
		return driver.findElement(backToSupport);
	}

	
	public WebElement GetSupportToolIndex(String indexName) {
		return driver.findElement(By.xpath("//a[text()='"+indexName+"']"));
	}
	
	public WebElement GetResultTable() {
		return driver.findElement(resultTable);
	}
	
	
	public void SelectValueFlowAction(String IndexName,String value,int row) throws Exception {
		GetSupportToolIndex(IndexName).click();
		new CommonLib().SelectDropDownValue(GetFlowAction(), value);
		GetListCampaign().click();
		new CommonLib().WriteExcelData("Sheet1", row, 0, value);
		new CommonLib().WriteExcelData("Sheet1", row, 1, GetCount());
		
	}
	
	public int GetCount() {
		int count = driver.findElements(actionCount).size();
		return count-1;
		
	}
	public void ClickBackToSupportLink() {
		GetBackToSupportLink().click();
	}
	
	

	

}
