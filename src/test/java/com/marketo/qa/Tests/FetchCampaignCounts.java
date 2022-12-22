package com.marketo.qa.Tests;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.Listeners;
import org.testng.annotations.Test;

import com.marketo.qa.FileLib.ListImpClass;
import com.marketo.qa.Pages.AdminPage;
import com.marketo.qa.Pages.MarketingActivitePage;
import com.marketo.qa.Pages.MyMarketoPage;
import com.marketo.qa.base.TestBase;

@Listeners(ListImpClass.class)
public class FetchCampaignCounts extends TestBase {
	MyMarketoPage homePage = new MyMarketoPage();
	MarketingActivitePage mAP = new MarketingActivitePage();
	private static Logger logger = LogManager.getLogger(TestBase.class);
	AdminPage admin = new AdminPage();

	@BeforeMethod
	public void PreRequest() throws Throwable {
		homePage.OpenAdminTab();
		try {
			admin.switchFrame();
			admin.GetTreasureChest().click();
			if (admin.GetCIstatus().getText().equalsIgnoreCase("Enabled")) {
				driver.switchTo().defaultContent();
				homePage.OpenMyMarketo();
			} else {
				admin.GetCIedit().click();
				admin.GetEnabledCheckbox().click();
				admin.GetCIsaveBtn().click();
				driver.switchTo().defaultContent();
				homePage.OpenMyMarketo();
			}
		} catch (Exception e) {
			driver.switchTo().defaultContent();
			logger.info("Treasure Chest Edit is not active for this user");
		}

	}

	@Test
	public void CampaignsCount() throws Throwable {

		homePage.OpenMarketingActivitiesTab();
		logger.info("Marketing Activite Page Task Opened");

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
		logger.info("Marketing Activite Page Task is done");

	}

}
