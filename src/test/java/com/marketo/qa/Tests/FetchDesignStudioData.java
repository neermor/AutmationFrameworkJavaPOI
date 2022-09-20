package com.marketo.qa.Tests;

import org.testng.annotations.Listeners;
import org.testng.annotations.Test;

import com.marketo.qa.FileLib.ListImpClass;
import com.marketo.qa.Pages.DesignStudioPage;
import com.marketo.qa.Pages.MarketingActivitePage;
import com.marketo.qa.Pages.MyMarketoPage;
import com.marketo.qa.base.TestBase;

@Listeners(ListImpClass.class)
public class FetchDesignStudioData extends TestBase {

	MyMarketoPage homePage = new MyMarketoPage();
	DesignStudioPage Ds = new DesignStudioPage();
	MarketingActivitePage mAP = new MarketingActivitePage();

	@Test
	public void TakeRequiredCount() throws Throwable {

		homePage.OpenDesignStudioTab();
		String WorkspaceCondition = prop.getProperty("WorkspaceCondition");

		switch (WorkspaceCondition) {
		case "All":
			Ds.AllWorkspaceRequiredCount();
			break;

		case "Specific":
			int NoOfWorkspace = Integer.parseInt(prop.getProperty("NoOfWorkspaces"));
			Ds.SpecificWorkspaceRequiredCount(NoOfWorkspace);
			break;

		default:
			break;
		}
	}

}
