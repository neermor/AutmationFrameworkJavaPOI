package com.marketo.qa.Tests;

import org.testng.annotations.Listeners;
import org.testng.annotations.Test;

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

	@Test
	public void ChampiensCount() throws Throwable {

		homePage.OpenMarketingActivitiesTab();
		mAP.GetEventCount(27);
		mAP.GetNurtureCount(28);
		mAP.GetProgrameLibrary(29);

		homePage.OpenMyMarketo();
		Wp.WebCampaignsScreenShot(32);
		driver.switchTo().defaultContent();

		homePage.OpenMyMarketo();
		mAP.switchFrame();
		Tam.VerifyAndFetchScreenshot(30);
		driver.switchTo().defaultContent();

		homePage.OpenMyMarketo();
		mAP.switchFrame();
		Pc.CheckPresentOrNot(31);
		driver.switchTo().defaultContent();

	}

}
