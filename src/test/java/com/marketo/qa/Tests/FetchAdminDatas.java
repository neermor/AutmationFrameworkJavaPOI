package com.marketo.qa.Tests;


import org.testng.annotations.Test;


import com.marketo.qa.Pages.AdminPage;
import com.marketo.qa.Pages.MarketingActivitePage;
import com.marketo.qa.Pages.MyMarketoPage;
import com.marketo.qa.base.TestBase;




public class FetchAdminDatas extends TestBase {
	MyMarketoPage homePage= new MyMarketoPage();
	AdminPage admin= new AdminPage();
	MarketingActivitePage mAP= new MarketingActivitePage();

	@Test	
	public void VerifyCountAndScreenshor()throws Throwable{
	MarketoLogin();
	homePage.OpenAdminTab();
	admin.switchFrame();
	admin.TagsCountAndScreenshot("Tags",18);
	admin.OpenLunchPoint(19);
	admin.GetInterestingMomentSubscription(20);
	driver.switchTo().defaultContent();
	Logout();
	}
	
	
	
}
