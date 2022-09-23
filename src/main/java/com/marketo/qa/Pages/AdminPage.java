package com.marketo.qa.Pages;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;

import com.marketo.qa.FileLib.CommonLib;
import com.marketo.qa.base.TestBase;
import com.marketo.qa.utility.screenshotUtility;

public class AdminPage extends TestBase {

	private static Logger logger = LogManager.getLogger(AdminPage.class);

	By MyAccount = By.xpath("//li[@class='x-tree-node']//span[text()='My Account']");
	By AccountName = By.xpath(
			"//span[text()='Account Settings']/../following-sibling::div//label[text()='Name:']/following-sibling::div//div");
	By Tags = By.xpath("//li[@class='x-tree-node']//span[text()='Tags']");
	By Table = By.cssSelector(
			"[class='x-panel-body x-panel-body-noheader x-panel-body-noborder'] div div div + [class='x-grid3-scroller'] [class='x-grid3-body']");
	By ChannelTag = By.xpath("//span [text()='Channel']/preceding-sibling:: img");
	By Iframe = By.cssSelector("#mlm");
	By TagCount = By.xpath("//span[text()='Channel']/../../../following-sibling:: td //div[@class='undefined'] /span");
	By IntegrationTag = By.xpath(
			"//span [text()='Integration']/../preceding-sibling:: img[@class='x-tree-ec-icon x-tree-elbow-minus']");
	By LunchPoint = By.xpath("//li[@class='x-tree-node']//span[text()='LaunchPoint']");
	By VerifyLunchPoint = By.xpath("//span[text()='No services configured']");
	By InstalledService = By.cssSelector("[class='x-grid3-viewport']");
	By LeadsCount = By.cssSelector("[class='x-toolbar-right-row'] [class='xtb-text']");
	By SalesInsight = By.xpath("//li[@class='x-tree-node']//span[text()='Sales Insight']");

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

	public WebElement GetIntegrationTag() {
		return driver.findElement(IntegrationTag);
	}

	public WebElement GetLunchPoint() {
		return driver.findElement(LunchPoint);
	}

	public WebElement GetVerifyLunchPoint() {
		return driver.findElement(VerifyLunchPoint);
	}

	public WebElement GetInstalledService() {
		return driver.findElement(InstalledService);
	}

	public void switchFrame() throws Throwable {
		driver.switchTo().defaultContent();
		driver.switchTo().frame(GetIFrame());

	}

	public void AccountName(int row) throws Throwable {
		new CommonLib().ClearExcelData("Sheet1", row);
		GetMyAccount().click();
		new CommonLib().WriteExcelData("Sheet1", row, 0, "Account Name");
		new CommonLib().WriteExcelData("Sheet1", row, 1, GetAccountName().getText());
		logger.info("Read Account Name");

	}

	public void TagsCountAndScreenshot(String tags, int row) throws Throwable {
		new CommonLib().ClearExcelData("Sheet1", row);
		new CommonLib().StandardWait(4000);
		GetTags().click();
		GetChannelTag().click();
		screenshotUtility.TakeScreenshot(GetTable(), "Tags1");
		logger.info("Click Tags Screenshots");
		new CommonLib().WriteExcelData("Sheet1", row, 0, tags);
		new CommonLib().WriteExcelData("Sheet1", row, 1, GetTagCount().getText());
		logger.info("Read Tags Count");

	}

	public void OpenLunchPoint(int row) throws Throwable {
		new CommonLib().ClearExcelData("Sheet1", row);

		GetLunchPoint().click();
		try {
			driver.findElement(By.xpath("//span[text()='No services configured']")).isDisplayed();
			logger.info("Integration not available");

		} catch (Exception e) {
			screenshotUtility.TakeScreenshot(GetInstalledService(), "Integration");
			logger.info("Integration Is available");
			logger.info("Click Integration Screenshot");
		}

		new CommonLib().WriteExcelData("Sheet1", row, 0, "Integration");
		new CommonLib().WriteExcelData("Sheet1", row, 1, GetCount());
	}

	public String GetCount() throws Throwable {

		Thread.sleep(4000);
		String countString = driver.findElement(LeadsCount).getText();
		String[] words = countString.split("\\s");
		if (words[0].equalsIgnoreCase("No")) {
			return "0";
		} else {
			return words[0];
		}
	}

	public void GetInterestingMomentSubscription(int row) throws Throwable {
		new CommonLib().ClearExcelData("Sheet1", row);

		try {
			driver.findElement(SalesInsight).isDisplayed();
			new CommonLib().WriteExcelData("Sheet1", row, 0, "Interesting Moment Subscription");
			new CommonLib().WriteExcelData("Sheet1", row, 1, "True");

		} catch (Exception e) {
			new CommonLib().WriteExcelData("Sheet1", row, 0, "Interesting Moment Subscription");
			new CommonLib().WriteExcelData("Sheet1", row, 1, "False");
		}
	}
}
