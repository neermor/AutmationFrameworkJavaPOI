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
			GhostLogin();
			homePage.VerifyHomeTileElements();
			homePage.OpenMarketingActivitiesTab();
			mAP.SelectWorkSpace("Default");
			mAP.switchFrame();
			mAP.GetCampaignCount("All Campaigns",7);
			mAP.GetCampaignCount("Active Campaigns", 8);
			mAP.GetMoreCampaignCount("Active Triggered Campaigns", 9);
			mAP.GetMoreCampaignCount("All Triggered Campaigns", 10);
			
	}


}
