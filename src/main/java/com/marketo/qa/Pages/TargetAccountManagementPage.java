package com.marketo.qa.Pages;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;

import com.marketo.qa.FileLib.CommonLib;
import com.marketo.qa.base.TestBase;
import com.marketo.qa.utility.screenshotUtility;

public class TargetAccountManagementPage extends TestBase {
	MyMarketoPage homepage = new MyMarketoPage();
	private static Logger logger = LogManager.getLogger(TestBase.class);
	CommonLib Clib = new CommonLib();

	By TopNamedAccounts = By.xpath("//div[@class='x4-surface x4-box-item x4-surface-default']");
	By NamedAccountNo = By.xpath("//label[text()='Named Accounts']/preceding-sibling::label");
	By Pipeline = By.xpath("//label[text()='Pipeline']/preceding-sibling::label");
	By Overview = By.xpath("//label[text()='Named Accounts']/../../../..");
	By OpenOpportunities = By.xpath(
			"//label[text()='Named Accounts']/../../../following-sibling::div//label[text()='Open Opportunities']/preceding-sibling::label");

	public WebElement GetTopNamedAccounts() {
		return driver.findElement(TopNamedAccounts);
	}

	public WebElement GetOverview() {
		return driver.findElement(Overview);
	}

	public WebElement GetOpenOpportunities() {
		return driver.findElement(OpenOpportunities);
	}

	public WebElement GetPipeline() {
		return driver.findElement(Pipeline);
	}

	public WebElement GetNamedAccountNo() {
		return driver.findElement(NamedAccountNo);
	}

	public void VerifyAndFetchScreenshot(int row, int NamedAccountsRow, int PipelineRow, int OpenOpportunitiesRow)
			throws Throwable {
		Clib.ClearExcelData("Sheet1", row);
		Clib.ClearExcelData("Sheet1", NamedAccountsRow);
		Clib.ClearExcelData("Sheet1", PipelineRow);
		Clib.ClearExcelData("Sheet1", OpenOpportunitiesRow);

		try {

			homepage.GetHometileUnderFrame("Target Account Management").click();
			logger.info("Target Account Management available");
			Thread.sleep(4000);

			Actions actions = new Actions(driver);
			actions.sendKeys(Keys.ESCAPE).perform();
			int NamedAccounts = Integer.parseInt(GetNamedAccountNo().getText());
			if (NamedAccounts > 0) {
				Clib.WriteExcelData("Sheet1", row, 0, "Target Account Management");
				Clib.WriteExcelData("Sheet1", row, 1, "True");
				new CommonLib().WriteExcelData("Sheet1", NamedAccountsRow, 0, "Named Accounts");
				new CommonLib().WriteExcelData("Sheet1", NamedAccountsRow, 1, NamedAccounts);
				new CommonLib().WriteExcelData("Sheet1", PipelineRow, 0, "Pipeline");
				new CommonLib().WriteExcelData("Sheet1", PipelineRow, 1, GetPipeline().getText());
				new CommonLib().WriteExcelData("Sheet1", OpenOpportunitiesRow, 0, "Open Opportunities");
				new CommonLib().WriteExcelData("Sheet1", OpenOpportunitiesRow, 1, GetOpenOpportunities().getText());

				Clib.StandardWait(4000);
				screenshotUtility.TakeScreenshot(GetTopNamedAccounts(), "Top Named Accounts");
				screenshotUtility.TakeScreenshot(GetOverview(), "Overview");

			} else {
				new CommonLib().WriteExcelData("Sheet1", NamedAccountsRow, 0, "Named Accounts");
				new CommonLib().WriteExcelData("Sheet1", NamedAccountsRow, 1, NamedAccounts);
				new CommonLib().WriteExcelData("Sheet1", PipelineRow, 0, "Pipeline");
				new CommonLib().WriteExcelData("Sheet1", PipelineRow, 1, GetPipeline().getText());
				new CommonLib().WriteExcelData("Sheet1", OpenOpportunitiesRow, 0, "Open Opportunities");
				new CommonLib().WriteExcelData("Sheet1", OpenOpportunitiesRow, 1, GetOpenOpportunities().getText());

				Clib.WriteExcelData("Sheet1", row, 0, "Target Account Management");
				Clib.WriteExcelData("Sheet1", row, 1, "True");
				logger.info("Target Account Management is Zero Count");
			}

		} catch (Exception e) {
			Clib.WriteExcelData("Sheet1", row, 0, "Target Account Management");
			Clib.WriteExcelData("Sheet1", row, 1, "False");
			logger.info("Target Account Management not available");

		}

	}

}
