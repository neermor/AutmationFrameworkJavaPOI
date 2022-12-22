package com.marketo.qa.Tests;

import org.testng.annotations.Test;

import com.marketo.qa.FileLib.CommonLib;
import com.marketo.qa.Pages.MarketingActivitePage;
import com.marketo.qa.Pages.MyMarketoPage;
import com.marketo.qa.Pages.PredictiveContentPage;
import com.marketo.qa.Pages.TargetAccountManagementPage;
import com.marketo.qa.Pages.WebPersonalization;
import com.marketo.qa.base.TestBase;

public class AddOnFeature extends TestBase {
	MyMarketoPage homePage = new MyMarketoPage();
	MarketingActivitePage mAP = new MarketingActivitePage();
	TargetAccountManagementPage Tam = new TargetAccountManagementPage();
	PredictiveContentPage Pc = new PredictiveContentPage();
	WebPersonalization Wp = new WebPersonalization();

	@Test
	public void AddOnFeatures() throws Throwable {

		new CommonLib().ClearExcelData("Sheet1", 34);
		new CommonLib().ClearExcelData("Sheet1", 35);
		new CommonLib().ClearExcelData("Sheet1", 36);
		new CommonLib().ClearExcelData("Sheet1", 37);
		new CommonLib().ClearExcelData("Sheet1", 38);
		new CommonLib().ClearExcelData("Sheet1", 39);
		new CommonLib().ClearExcelData("Sheet1", 40);
		new CommonLib().ClearExcelData("Sheet1", 41);

		homePage.OpenMyMarketo();
		mAP.switchFrame();
		Tam.VerifyAndFetchScreenshot(34, 35, 36, 37);
		driver.switchTo().defaultContent();

		homePage.OpenMyMarketo();
		mAP.switchFrame();
		Pc.CheckPresentOrNot(38, 39);
		driver.switchTo().defaultContent();

		homePage.OpenMyMarketo();
		Wp.WebCampaignsScreenShot(40, prop.getProperty("WorkspaceCondition"), 41);
		driver.switchTo().defaultContent();

	}

}
