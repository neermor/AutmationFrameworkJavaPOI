package com.marketo.qa.Tests;

import org.testng.annotations.Listeners;
import org.testng.annotations.Test;

import com.marketo.qa.FileLib.ListImpClass;
import com.marketo.qa.Pages.AdminPage;
import com.marketo.qa.Pages.MarketingActivitePage;
import com.marketo.qa.Pages.MyMarketoPage;
import com.marketo.qa.base.TestBase;

@Listeners(ListImpClass.class)
public class FetchAdminDatas extends TestBase {
	MyMarketoPage homePage = new MyMarketoPage();
	AdminPage admin = new AdminPage();
	MarketingActivitePage mAP = new MarketingActivitePage();

	@Test
	public void VerifyCountAndScreenshor() throws Throwable {
		homePage.OpenAdminTab();
		admin.switchFrame();
		admin.TagsCountAndScreenshot("Tags", 19);
		admin.OpenLunchPoint(20);
		admin.GetInterestingMomentSubscription(21);
		admin.AccountName(22);

		driver.switchTo().defaultContent();
	}

}
