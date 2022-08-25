package com.marketo.qa.Tests;

import org.testng.annotations.Test;

import com.marketo.qa.Pages.DesignStudioPage;
import com.marketo.qa.Pages.MarketingActivitePage;
import com.marketo.qa.Pages.MyMarketoPage;
import com.marketo.qa.base.TestBase;

public class FetchDesignStudioData extends TestBase{
	
	MyMarketoPage homePage= new MyMarketoPage();
	DesignStudioPage Ds= new DesignStudioPage();
	MarketingActivitePage mAP= new MarketingActivitePage();


	@Test
	public void TakeRequiredCount() throws Throwable {
		
		homePage.OpenDesignStudioTab();	
		Ds.AllWorkspaceRequiredCount();

		
	}

}
