package com.marketo.qa.Tests;

import org.testng.annotations.Test;

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

		homePage.OpenMyMarketo();
		mAP.switchFrame();
		Tam.VerifyAndFetchScreenshot(34);
		driver.switchTo().defaultContent();

		homePage.OpenMyMarketo();
		mAP.switchFrame();
		Pc.CheckPresentOrNot(35);
		driver.switchTo().defaultContent();

		homePage.OpenMyMarketo();
		Wp.WebCampaignsScreenShot(36, prop.getProperty("WorkspaceCondition"));
		driver.switchTo().defaultContent();

	}

}
