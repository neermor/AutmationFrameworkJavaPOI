package com.marketo.qa.Tests;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.testng.annotations.Listeners;
import org.testng.annotations.Test;

import com.marketo.qa.FileLib.CommonLib;
import com.marketo.qa.FileLib.ListImpClass;
import com.marketo.qa.Pages.MarketingActivitePage;
import com.marketo.qa.Pages.MyMarketoPage;
import com.marketo.qa.Pages.PredictiveContentPage;
import com.marketo.qa.Pages.TargetAccountManagementPage;
import com.marketo.qa.Pages.WebPersonalization;
import com.marketo.qa.base.TestBase;

@Listeners(ListImpClass.class)
public class FetchNuture_EventCount extends TestBase {
	MyMarketoPage homePage = new MyMarketoPage();
	MarketingActivitePage mAP = new MarketingActivitePage();
	TargetAccountManagementPage Tam = new TargetAccountManagementPage();
	PredictiveContentPage Pc = new PredictiveContentPage();
	WebPersonalization Wp = new WebPersonalization();
	private static Logger logger = LogManager.getLogger(TestBase.class);

	@Test
	public void ChampiensCount() throws Throwable {

		new CommonLib().ClearExcelData("Sheet1", 31);
		new CommonLib().ClearExcelData("Sheet1", 32);
		new CommonLib().ClearExcelData("Sheet1", 33);

		homePage.OpenMarketingActivitiesTab();
		logger.info("Marketing Activite Page Task Opened");
		new CommonLib().WaitForElementToLoad(driver, 60, mAP.GetFilter());
		mAP.GetFilter().click();
		mAP.GetEventCount(31);
		mAP.GetNurtureCount(32);
		mAP.GetProgramLibrary(33);

		logger.info("Marketing Activite Page Task Closed");

	}

}
