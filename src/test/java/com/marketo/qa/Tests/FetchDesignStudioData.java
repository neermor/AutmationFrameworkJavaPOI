package com.marketo.qa.Tests;

import org.testng.annotations.Test;

import com.marketo.qa.Pages.DesignStudioPage;
import com.marketo.qa.Pages.MyMarketoPage;
import com.marketo.qa.base.TestBase;

public class FetchDesignStudioData extends TestBase{
	
	MyMarketoPage homePage= new MyMarketoPage();
	DesignStudioPage Ds= new DesignStudioPage();

	@Test
	public void TakeRequiredCount() throws Throwable {
		
		MarketoLogin();
		homePage.VerifyHomeTileElements();
		homePage.OpendesignStudioTab();		
		Ds.FetchTreeNodeCount("Emails",4);
		Ds.FetchTreeNodeCount("Forms",5);
		Ds.FetchTreeNodeCount("Landing Pages",6);
		Ds.FetchUploadCount("Images and Files",7);
		Ds.FetchSnippetsCount("Snippets",8);		
		
	}

}
