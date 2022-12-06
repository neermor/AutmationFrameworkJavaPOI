package com.marketo.qa.Tests;

import org.testng.annotations.Listeners;
import org.testng.annotations.Test;

import com.marketo.qa.FileLib.CommonLib;
import com.marketo.qa.FileLib.ListImpClass;
import com.marketo.qa.Pages.MarketingActivitePage;
import com.marketo.qa.Pages.MyMarketoPage;
import com.marketo.qa.Pages.SupportToolsPage;
import com.marketo.qa.base.TestBase;
import com.marketo.qa.utility.screenshotUtility;

@Listeners(ListImpClass.class)
public class FetchDataFromChangeScore extends TestBase {
	MyMarketoPage homePage = new MyMarketoPage();
	MarketingActivitePage mAP = new MarketingActivitePage();
	SupportToolsPage Support = new SupportToolsPage();

	@Test
	public void VerifyChangeScoreCount() throws Throwable {
		new CommonLib().ClearExcelData("Sheet1", 25);
		new CommonLib().ClearExcelData("Sheet1", 26);
		new CommonLib().ClearExcelData("Sheet1", 27);

		OpenSupportTool();
		Support.SelectValueFlowAction("Flow Actions Used", "Change Score", 25);
		screenshotUtility.TakeScreenshot(Support.GetResultTable(), "Change Score");
		Support.ClickBackToSupportLink();
		Support.SelectValueFlowAction("Flow Actions Used", "Interesting Moment", 26);
		screenshotUtility.TakeScreenshot(Support.GetResultTable(), "Interesting Moment");
		Support.ClickBackToSupportLink();
		Support.SelectValueFlowAction("Flow Actions Used", "Change Data Value", 27);
		screenshotUtility.TakeScreenshot(Support.GetResultTable(), "Change Data Value");

	}

}
