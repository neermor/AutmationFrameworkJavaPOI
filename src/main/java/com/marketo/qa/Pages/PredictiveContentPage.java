package com.marketo.qa.Pages;

import java.util.Iterator;
import java.util.Set;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;

import com.marketo.qa.FileLib.CommonLib;
import com.marketo.qa.base.TestBase;
import com.marketo.qa.utility.screenshotUtility;

public class PredictiveContentPage extends TestBase {
	MyMarketoPage homepage = new MyMarketoPage();
	private static Logger logger = LogManager.getLogger(TestBase.class);
	String Parent_window = null;

	By TotalContentNo = By.xpath("//div[text()='TOTAL CONTENT']/preceding-sibling::div");

	public WebElement GetTotalContentNo() {
		return driver.findElement(TotalContentNo);
	}

	public void CheckPresentOrNot(int row, int TCRow) throws Throwable {
		try {
			homepage.GetHometileUnderFrame("Predictive Content").click();
			logger.info("Predictive Content available");
			Set<String> s = driver.getWindowHandles();
			Iterator<String> I1 = s.iterator();

			Parent_window = I1.next();
			String child_window = I1.next();
			driver.switchTo().window(child_window);
			new CommonLib().StandardWait(6000);
			int TC = Integer.parseInt(GetTotalContentNo().getText());
			if (TC > 0) {
				screenshotUtility.TakeFullPageScreenshot("Predictive Content");
				new CommonLib().WriteExcelData("Sheet1", row, 0, "Predictive Content");
				new CommonLib().WriteExcelData("Sheet1", row, 1, "True");
				new CommonLib().WriteExcelData("Sheet1", TCRow, 0, "Total Content");
				new CommonLib().WriteExcelData("Sheet1", TCRow, 1, TC);

				driver.close();
				driver.switchTo().window(Parent_window);
			} else {
				new CommonLib().WriteExcelData("Sheet1", row, 0, "Predictive Content");
				new CommonLib().WriteExcelData("Sheet1", row, 1, "True");
				new CommonLib().WriteExcelData("Sheet1", TCRow, 0, "Total Content");
				new CommonLib().WriteExcelData("Sheet1", TCRow, 1, TC);

				driver.close();
				driver.switchTo().window(Parent_window);

				logger.info("Predictive Content is Zero Count");

			}

		}

		catch (Exception e) {
			new CommonLib().WriteExcelData("Sheet1", row, 0, "Predictive Content");
			new CommonLib().WriteExcelData("Sheet1", row, 1, "False");
			logger.info("Predictive Content not available");

		}

	}
}
