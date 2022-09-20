package com.marketo.qa.Tests;

import org.testng.annotations.Listeners;
import org.testng.annotations.Test;

import com.marketo.qa.FileLib.ListImpClass;
import com.marketo.qa.Pages.MarketingActivitePage;
import com.marketo.qa.Pages.MyMarketoPage;
import com.marketo.qa.base.TestBase;

@Listeners(ListImpClass.class)
public class FetchChampianCounts extends TestBase {
	MyMarketoPage homePage = new MyMarketoPage();
	MarketingActivitePage mAP = new MarketingActivitePage();

	@Test
	public void ChampiensCount() throws Throwable {
		homePage.OpenMarketingActivitiesTab();
		String WorkspaceCondition = prop.getProperty("WorkspaceCondition");

		switch (WorkspaceCondition) {
		case "All":
			mAP.AllWorkspaceCampaignCount();
			break;

		case "Specific":
			int NoOfWorkspace = Integer.parseInt(prop.getProperty("NoOfWorkspaces"));
			mAP.SpecificWorkspaceCampaignCount(NoOfWorkspace);
			break;

		default:
			break;
		}
	}

}
