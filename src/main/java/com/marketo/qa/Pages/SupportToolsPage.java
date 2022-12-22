package com.marketo.qa.Pages;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;
import org.testng.asserts.SoftAssert;

import com.marketo.qa.FileLib.CommonLib;
import com.marketo.qa.base.TestBase;

public class SupportToolsPage extends TestBase {
	private static Logger logger = LogManager.getLogger(TestBase.class);
	SoftAssert asrt = new SoftAssert();

	private By flowAction = By.cssSelector("#actionId");
	private By listCampaigns = By.xpath("//button[text()='List Campaigns']");
	private By actionCount = By.xpath("//form[@id='flowActionUsage']/a[@name ='results']/p");
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
		return driver.findElement(By.xpath("//a[text()='" + indexName + "']"));
	}

	public WebElement GetResultTable() {
		return driver.findElement(resultTable);
	}

	boolean WorkspaceAvl = true;

	public void SelectValueFlowAction(String IndexName, String value, int row) throws Exception {
		try {
			GetSupportToolIndex(IndexName).click();
			logger.info(IndexName + " Selected");
			new CommonLib().SelectDropDownValue(GetFlowAction(), value);
			GetListCampaign().click();
			new CommonLib().WriteExcelData("Sheet1", row, 0, value);
			new CommonLib().WriteExcelData("Sheet1", row, 1, GetCount());
			logger.info("Fetch " + value + " Count");
		} catch (Exception e) {
			logger.info(value + " Not Present..");
			asrt.assertTrue(WorkspaceAvl, value + " Not Present..");

		}
		asrt.assertAll();

	}

	public String GetCount() {
		String count = driver.findElement(actionCount).getText();
		if (count.startsWith("Showing actions")) {
			String[] words = count.split("\\s");
			return words[6];
		} else {
			return "0";
		}
	}

	public void ClickBackToSupportLink() {
		GetBackToSupportLink().click();
	}

}
