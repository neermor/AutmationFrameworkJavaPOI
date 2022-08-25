package com.marketo.qa.Tests;

import org.testng.annotations.Test;

import com.marketo.qa.Pages.MarketingActivitePage;
import com.marketo.qa.Pages.MyMarketoPage;
import com.marketo.qa.base.TestBase;


public class FetchChampianCounts extends TestBase {
	MyMarketoPage homePage= new MyMarketoPage();
	MarketingActivitePage mAP= new MarketingActivitePage();
	
	@Test
	public void ChampiensCount() throws Throwable {
			homePage.OpenMarketingActivitiesTab();
			mAP.AllWorkspaceCampaignCount();

	}



}
