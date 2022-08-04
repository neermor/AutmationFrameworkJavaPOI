package com.marketo.qa.Tests;


import org.testng.annotations.Test;

import com.marketo.qa.Pages.AdminPage;
import com.marketo.qa.Pages.MyMarketoPage;
import com.marketo.qa.base.TestBase;

public class FetchAdminDatas extends TestBase {
	MyMarketoPage homePage= new MyMarketoPage();
	AdminPage admin= new AdminPage();
	@Test	
	public void VerifyCountAndScreenshor()throws Throwable{
	MarketoLogin();
	homePage.OpenAdminTab();
	Thread.sleep(10000);
	admin.switchFrame();
	admin.TagsCountAndScreenshot("Tags",11);
	}
	
}
