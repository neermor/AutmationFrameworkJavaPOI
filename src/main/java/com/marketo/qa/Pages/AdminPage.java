package com.marketo.qa.Pages;

import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;

import com.marketo.qa.FileLib.CommonLib;
import com.marketo.qa.base.TestBase;
import com.marketo.qa.utility.screenshotUtility;

public class AdminPage extends TestBase{
	
	By MyAccount=By.xpath("//li[@class='x-tree-node']//span[text()='My Account']");
	By AccountName = By.xpath("//span[text()='Account Settings']/../following-sibling::div//label[text()='Name:']/following-sibling::div//div");
	By Tags=By.xpath("//li[@class='x-tree-node']//span[text()='Tags']");
	By Table = By.cssSelector("[class='x-panel-body x-panel-body-noheader x-panel-body-noborder']");
	By ChannelTag = By.xpath("//span [text()='Channel']/preceding-sibling:: img");
	By Iframe = By.cssSelector("#mlm");
	By TagCount = By.xpath("//span[text()='Channel']/../../../following-sibling:: td //div[@class='undefined'] /span");
	
	public WebElement GetMyAccount() {
		return driver.findElement(MyAccount);
	}
	
	public WebElement GetTagCount() {
		return driver.findElement(TagCount);
	}
	
	public WebElement GetAccountName() {
		return driver.findElement(AccountName);
	}
	
	public WebElement GetTags() {
		return driver.findElement(Tags);
	}
	
	public WebElement GetTable() {
		return driver.findElement(Table);
	}
	
	public WebElement GetChannelTag() {
		return driver.findElement(ChannelTag);
	}
	public WebElement GetIFrame() {
		return driver.findElement(Iframe);
	}
	
	public void switchFrame() throws Throwable {
		driver.switchTo().frame(GetIFrame());

	}
	
	public String AccountName() {
		GetMyAccount().click();
		return GetAccountName().getText();
		
	}
	
	public void TagsCountAndScreenshot(String tags, int row) throws Throwable {
		GetTags().click();
		GetChannelTag().click();
		screenshotUtility.TakeScreenshot(GetTable(), tags);
		new CommonLib().WriteExcelData("Sheet1", row, 0, tags);
		new CommonLib().WriteExcelData("Sheet1", row, 1, GetTagCount().getText());
		
	}
	
	
	

	

}
