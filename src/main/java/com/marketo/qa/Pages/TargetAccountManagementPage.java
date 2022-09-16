package com.marketo.qa.Pages;

import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;

import com.marketo.qa.FileLib.CommonLib;
import com.marketo.qa.base.TestBase;
import com.marketo.qa.utility.screenshotUtility;

public class TargetAccountManagementPage extends TestBase {
	MyMarketoPage homepage = new MyMarketoPage();

	By Overview = By.xpath("//div[contains(@class,'abmDashboard') and contains(@id,'abmDashboard')]");

	public WebElement GetOverview() {
		return driver.findElement(Overview);
	}

	public void VerifyAndFetchScreenshot(int row) throws Throwable {
		new CommonLib().ClearExcelData("Sheet1", row);

		try {

			homepage.GetHometileUnderFrame("Target Account Management").click();
			new CommonLib().WriteExcelData("Sheet1", row, 0, "Target Account Management");
			new CommonLib().WriteExcelData("Sheet1", row, 1, "True");
			new CommonLib().StandardWait(4000);
			screenshotUtility.TakeScreenshot(GetOverview(), "Target Account Management");

		} catch (Exception e) {
			new CommonLib().WriteExcelData("Sheet1", row, 0, "Target Account Management");
			new CommonLib().WriteExcelData("Sheet1", row, 1, "False");
		}

	}

}
