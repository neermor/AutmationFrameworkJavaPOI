package com.marketo.qa.Tests;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.testng.annotations.Listeners;
import org.testng.annotations.Test;

import com.marketo.qa.FileLib.ListImpClass;
import com.marketo.qa.Pages.DatabasePage;
import com.marketo.qa.Pages.MarketingActivitePage;
import com.marketo.qa.Pages.MyMarketoPage;
import com.marketo.qa.base.TestBase;

@Listeners(ListImpClass.class)
public class FetchLeadsCount extends TestBase {
	MyMarketoPage homePage = new MyMarketoPage();
	DatabasePage Db = new DatabasePage();
	MarketingActivitePage mAP = new MarketingActivitePage();
	private static Logger logger = LogManager.getLogger(TestBase.class);

	@Test
	public void CollectLeadsCount() throws Throwable {
		homePage.OpenDatabaseTab();
		logger.info("Database Page Task Opened");

		String WorkspaceCondition = prop.getProperty("WorkspaceCondition");

		switch (WorkspaceCondition) {
		case "All":
			Db.AllWorkspaceCollectLeadsCount();
			break;

		case "Specific":
			int NoOfWorkspace = Integer.parseInt(prop.getProperty("NoOfWorkspaces"));
			Db.SpecificWorkspaceCollectLeadsCount(NoOfWorkspace);
			break;

		default:
			break;
		}
		logger.info("Database Page Task Closed");

	}

}
