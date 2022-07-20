package com.marketo.qa.Pages;

import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;

import com.marketo.qa.base.TestBase;

public class MarketingActivitesPage extends TestBase{

	private By marketingActivitesMenu=By.cssSelector("#mktExplorerFilterMenu > tbody > tr:nth-child(2) > td.x-btn-mc");
	private By AllCampaigns = By.xpath("//img[@class='x-menu-item-icon mkiLightbulb'] /following-sibling:: span[text()='All Campaigns']");
	
	public WebElement getMarketingActivitesMenu() {
		return driver.findElement(marketingActivitesMenu);
	}
	
	public WebElement getAllCampaigns() {
		return driver.findElement(AllCampaigns);
	}
	
	public void ClickMarketingActivitesMenu() throws Throwable {
		getMarketingActivitesMenu().click();
	}
	
	public void ClickAllCampaignsAndGetCount() {
		getAllCampaigns().click();
		
	}

	
}
