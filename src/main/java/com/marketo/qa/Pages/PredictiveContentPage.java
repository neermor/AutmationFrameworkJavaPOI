package com.marketo.qa.Pages;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import com.marketo.qa.FileLib.CommonLib;
import com.marketo.qa.base.TestBase;

public class PredictiveContentPage extends TestBase {
	MyMarketoPage homepage = new MyMarketoPage();
	private static Logger logger = LogManager.getLogger(TestBase.class);

	public void CheckPresentOrNot(int row) throws Exception {
		try {
			homepage.GetHometileUnderFrame("Predictive Content").isDisplayed();
			logger.info("Predictive Content available");
			new CommonLib().WriteExcelData("Sheet1", row, 0, "Predictive Content");
			new CommonLib().WriteExcelData("Sheet1", row, 1, "True");
		}

		catch (Exception e) {
			new CommonLib().WriteExcelData("Sheet1", row, 0, "Predictive Content");
			new CommonLib().WriteExcelData("Sheet1", row, 1, "False");
			logger.info("Predictive Content not available");

		}

	}
}
