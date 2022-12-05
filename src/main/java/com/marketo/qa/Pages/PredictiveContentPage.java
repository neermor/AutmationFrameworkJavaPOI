package com.marketo.qa.Pages;

import java.util.Iterator;
import java.util.Set;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import com.marketo.qa.FileLib.CommonLib;
import com.marketo.qa.base.TestBase;
import com.marketo.qa.utility.screenshotUtility;

public class PredictiveContentPage extends TestBase {
	MyMarketoPage homepage = new MyMarketoPage();
	private static Logger logger = LogManager.getLogger(TestBase.class);
	String Parent_window = null;

	public void CheckPresentOrNot(int row) throws Exception {
		try {
			homepage.GetHometileUnderFrame("Predictive Content").click();
			logger.info("Predictive Content available");
			Set<String> s = driver.getWindowHandles();
			Iterator<String> I1 = s.iterator();

			Parent_window = I1.next();
			String child_window = I1.next();
			driver.switchTo().window(child_window);

			screenshotUtility.TakeFullPageScreenshot("Predictive Content");
			new CommonLib().WriteExcelData("Sheet1", row, 0, "Predictive Content");
			new CommonLib().WriteExcelData("Sheet1", row, 1, "True");

			driver.close();
			driver.switchTo().window(Parent_window);
		}

		catch (Exception e) {
			new CommonLib().WriteExcelData("Sheet1", row, 0, "Predictive Content");
			new CommonLib().WriteExcelData("Sheet1", row, 1, "False");
			logger.info("Predictive Content not available");

		}

	}
}
