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
			homePage.OpenMarketingActivitiesTab();
			mAP.SelectWorkSpace("Default");
			mAP.switchFrame();
			mAP.GetMoreCampaignCount("All Triggered Campaigns",9);
			 mAP.GetMoreCampaignCount("Active Triggered Campaigns", 10);
			  mAP.GetMoreCampaignCount("Batch Campaigns - Repeating Schedule", 11);
			  mAP.GetMoreCampaignCount("All Batch Campaigns", 12);
			  mAP.GetCampaignCount("All Campaigns",13);
			  mAP.GetCampaignCount("Active Campaigns", 14);
			 driver.switchTo().defaultContent();			 

	}


}
