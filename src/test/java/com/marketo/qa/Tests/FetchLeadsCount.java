package com.marketo.qa.Tests;

import com.marketo.qa.Pages.MyMarketoPage;

import org.testng.annotations.Test;

import com.marketo.qa.Pages.DatabasePage;
import com.marketo.qa.Pages.MarketingActivitePage;
import com.marketo.qa.base.TestBase;

public class FetchLeadsCount extends TestBase {
	MyMarketoPage homePage= new MyMarketoPage();
	DatabasePage Db= new DatabasePage();

	@Test
	public void CollectLeadsCount() throws Throwable {
		MarketoLogin();
		homePage.OpenDatabaseTab();
		homePage.SelectWorkSpace("Default");
		Db.SegmentationsCount(17);
		Db.ExtendWorkshoptreenode("System Smart Lists","All People");
		Db.GetLeadsCount(15);
	}

}
