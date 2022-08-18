package com.marketo.qa.Tests;

import org.testng.annotations.Test;

import com.marketo.qa.Pages.AnalyticsPage;
import com.marketo.qa.Pages.MarketingActivitePage;
import com.marketo.qa.Pages.MyMarketoPage;
import com.marketo.qa.base.TestBase;

public class FetchModelsCount extends TestBase {
	MyMarketoPage homePage= new MyMarketoPage();
	AnalyticsPage Analytic= new AnalyticsPage();
	MarketingActivitePage mAP= new MarketingActivitePage();

	@Test
	public void ModelCount() throws Throwable {
		homePage.OpenanalyticsTab();
		homePage.SelectWorkSpace("Default");
		Analytic.ModelCount(16);
		}

}
