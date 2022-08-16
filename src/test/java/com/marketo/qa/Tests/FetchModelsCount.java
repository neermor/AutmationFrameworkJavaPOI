package com.marketo.qa.Tests;

import org.testng.annotations.Test;

import com.marketo.qa.Pages.AnalyticsPage;
import com.marketo.qa.Pages.MyMarketoPage;
import com.marketo.qa.base.TestBase;

public class FetchModelsCount extends TestBase {
	MyMarketoPage homePage= new MyMarketoPage();
	AnalyticsPage Analytic= new AnalyticsPage();
	
	@Test
	public void ModelCount() throws Throwable {
		MarketoLogin();
		homePage.OpenanalyticsTab();
		homePage.SelectWorkSpace("Default");
		Analytic.ModelCount(16);
		}

}
