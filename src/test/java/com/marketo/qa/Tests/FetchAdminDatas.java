package com.marketo.qa.Tests;

import org.testng.annotations.Listeners;
import org.testng.annotations.Test;

import com.marketo.qa.FileLib.CommonLib;
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
	public void VerifyCountAndScreenshot() throws Throwable {
		new CommonLib().ClearExcelData("Sheet1", 24);
		new CommonLib().ClearExcelData("Sheet1", 44);
		new CommonLib().ClearExcelData("Sheet1", 26);
		new CommonLib().ClearExcelData("Sheet1", 27);

		homePage.OpenAdminTab();
		admin.switchFrame();
		admin.TagsCountAndScreenshot("Tags", 24, 44);
		admin.OpenLunchPoint(25);
		admin.GetInterestingMomentSubscription(26);
		admin.AccountName(27);

		driver.switchTo().defaultContent();
	}

}
