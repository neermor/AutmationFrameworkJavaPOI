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

	By Overview = By.xpath("//div[contains(@class,'abmDashboard') and contains(@id,'abmDashboard')]");

	public WebElement GetOverview() {
		return driver.findElement(Overview);
	}

	public void VerifyAndFetchScreenshot(int row) throws Throwable {
		new CommonLib().ClearExcelData("Sheet1", row);

		try {

			homepage.GetHometileUnderFrame("Target Account Management").click();
			logger.info("Target Account Management available");
			Thread.sleep(4000);

			Actions actions = new Actions(driver);
			actions.sendKeys(Keys.ESCAPE).perform();
			new CommonLib().WriteExcelData("Sheet1", row, 0, "Target Account Management");
			new CommonLib().WriteExcelData("Sheet1", row, 1, "True");
			new CommonLib().StandardWait(8000);
			screenshotUtility.TakeFullPageScreenshot("Target Account Management");

		} catch (Exception e) {
			new CommonLib().WriteExcelData("Sheet1", row, 0, "Target Account Management");
			new CommonLib().WriteExcelData("Sheet1", row, 1, "False");
			logger.info("Target Account Management not available");

		}

	}

}
