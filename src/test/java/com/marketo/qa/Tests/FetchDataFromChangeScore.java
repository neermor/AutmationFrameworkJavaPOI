package com.marketo.qa.Tests;


import org.testng.annotations.Test;

import com.marketo.qa.Pages.MarketingActivitePage;
import com.marketo.qa.Pages.MyMarketoPage;
import com.marketo.qa.Pages.SupportToolsPage;
import com.marketo.qa.base.TestBase;
import com.marketo.qa.utility.screenshotUtility;

public class FetchDataFromChangeScore extends TestBase {
	MyMarketoPage homePage= new MyMarketoPage();
	MarketingActivitePage mAP= new MarketingActivitePage();
	SupportToolsPage Support = new SupportToolsPage();
	
	
	@Test
	public void VerifyChangeScoreCount() throws Throwable {
		OpenSupportTool();
		Support.SelectValueFlowAction("Flow Actions Used","Change Score",23);
		screenshotUtility.TakeScreenshot(Support.GetResultTable(), "Change Score");
		Support.ClickBackToSupportLink();
		Support.SelectValueFlowAction("Flow Actions Used","Interesting Moment",24);
		screenshotUtility.TakeScreenshot(Support.GetResultTable(),"Interesting Moment");
		Support.ClickBackToSupportLink();
		Support.SelectValueFlowAction("Flow Actions Used","Change Data Value",25);
		screenshotUtility.TakeScreenshot(Support.GetResultTable(), "Change Data Value");

		}
		
	
}
