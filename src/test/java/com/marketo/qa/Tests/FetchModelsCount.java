package com.marketo.qa.Tests;

import org.testng.annotations.Listeners;
import org.testng.annotations.Test;

import com.marketo.qa.FileLib.ListImpClass;
import com.marketo.qa.Pages.AnalyticsPage;
import com.marketo.qa.Pages.MyMarketoPage;
import com.marketo.qa.base.TestBase;

@Listeners(ListImpClass.class)
public class FetchModelsCount extends TestBase {
	MyMarketoPage homePage = new MyMarketoPage();
	AnalyticsPage Analytic = new AnalyticsPage();

	@Test
	public void ModelCount() throws Throwable {
		homePage.OpenanalyticsTab();
		String WorkspaceCondition = prop.getProperty("WorkspaceCondition");

		switch (WorkspaceCondition) {
		case "All":
			Analytic.AllWorkspaceModelCount();
			break;

		case "Specific":
			int NoOfWorkspace = Integer.parseInt(prop.getProperty("NoOfWorkspaces"));
			Analytic.SpecificWorkspaceModelCount(NoOfWorkspace);
			break;

		default:
			break;
		}
	}

}
