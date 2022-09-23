package com.marketo.qa.Pages;

import java.util.List;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;

import com.marketo.qa.FileLib.CommonLib;
import com.marketo.qa.base.TestBase;
import com.marketo.qa.utility.screenshotUtility;

public class MarketingActivitePage extends TestBase {
	private static Logger logger = LogManager.getLogger(TestBase.class);

	boolean flag = false;
	By GlobalTreeSearch = By.xpath("//input[@data-id='globalTreeSearchInput_input']");
	By TreeNode = By.xpath("//div[contains(@data-id,'treeNodeRow' )]");
	By WorkSpace = By.xpath("//div[contains(@data-id,'treeNode_workspace')]/..//div//span");
	By Iframe = By.cssSelector("#mlm");
	By CampaignInspector = By.cssSelector("#canvas__cp_campaignInspector");
	By CampaignDD = By.xpath("//button[text()='Active Campaigns']");
	By CampaignInspectorOption = By.cssSelector("[class='x-menu x-menu-floating x-layer'] ul li");
	By AllCount = By.cssSelector("[class='x-toolbar-right-ct'] [class='x-toolbar-cell'] div");
	By MoreCampaigns = By.xpath("//span[contains(text(),'More Campaigns')]");
	By MoreCampaignOptions = By.cssSelector("[class='x-menu x-menu-floating x-layer mktSubMenu'] ul li span");
	By filter = By.xpath("//button[@data-id='global_treeFilterButton']");
	By EventFilter = By.xpath("//input[@value='Event Programs']");
	By EngagementPrograms = By.xpath("//input[@value='Engagement Programs']");
	By resetBtn = By.cssSelector("[data-id='DrawerFooter'] button[data-id='globalTreeFilter_reset']");
	By NoResult = By.xpath("//div[@data-id='Tree_NoResultsText']");
	By Event = By.xpath("//div[contains(@data-id,'treeNode_eventprogram')]/..");
	By Nuture = By.xpath("//div[contains(@data-id,'treeNode_engagementprogram')]/..");
	By Page = By.xpath("//td[@class='x-toolbar-cell'] //div[contains(text(),'of ')]");
	By Template = By.xpath("//div[@data-id='Tree_TreeBodyLoadedFadedIn']");

	public WebElement GetTemplate() {
		return driver.findElement(Template);
	}

	public WebElement GetIFrame() {
		return driver.findElement(Iframe);
	}

	public WebElement GetGlobalTreeSearch() {
		return driver.findElement(GlobalTreeSearch);
	}

	public WebElement GetPage() {
		return driver.findElement(Page);
	}

	public WebElement GetResetBtn() {
		return driver.findElement(resetBtn);
	}

	public WebElement GetFilter() {
		return driver.findElement(filter);
	}

	public WebElement GetEventFilter() {
		return driver.findElement(EventFilter);
	}

	public WebElement GetEngagementPrograms() {
		return driver.findElement(EngagementPrograms);
	}

	public WebElement GetCampaignDD() {
		return driver.findElement(CampaignDD);
	}

	public WebElement GetCampaignInspector() {
		return driver.findElement(CampaignInspector);
	}

	public WebElement GetMoreCampaign() {
		return driver.findElement(MoreCampaigns);
	}

	public void switchFrame() throws Throwable {
		driver.switchTo().frame(GetIFrame());

	}

	public void ClickCampaignInspector() throws Throwable {
		new CommonLib().StandardWait(2000);
		GetCampaignInspector().click();
		GetCampaignDD().click();

	}

	public void SelectTreeNode(String TreeNodeName) throws Throwable {
		new CommonLib().StandardWait(2000);
		;
		List<WebElement> HomeTiles = driver.findElements(TreeNode);
		for (WebElement option : HomeTiles) {

			if (option.getText().equalsIgnoreCase(TreeNodeName)) {
				option.click();
			}
		}
	}

	public void GetMoreCampaignOption(String MoreCampaignValue) throws Throwable {
		Thread.sleep(4000);
		List<WebElement> CampaignOption = driver.findElements(MoreCampaignOptions);
		for (WebElement option : CampaignOption) {
			if (option.getText().equalsIgnoreCase(MoreCampaignValue)) {
				option.click();
			}
		}
	}

	public void GetCampaignInspectorOption(String CampaignType) throws Throwable {
		Thread.sleep(10000);
		List<WebElement> CampaignOption = driver.findElements(CampaignInspectorOption);
		for (WebElement option : CampaignOption) {

			if (option.getText().equalsIgnoreCase(CampaignType)) {
				option.click();
			}
		}
	}

	public void SelectWorkSpace(String WorkspaceName) throws Throwable {
		SelectTreeNode(WorkspaceName);
	}

	public String GetCount() throws Throwable {

		Thread.sleep(4000);
		String countString = driver.findElement(AllCount).getText();
		String[] words = countString.split("\\s");

		if (words[0].equalsIgnoreCase("No")) {
			return "0";
		} else {
			if (GetPage().getText().equalsIgnoreCase("Of 1")) {
				return words[0];

			} else {
				return words[2];
			}
		}
	}

	public int GetCampaignCount(String CampaignType, int row, int cell) throws Throwable {
		ClickCampaignInspector();
		GetCampaignInspectorOption(CampaignType);
		new CommonLib().WriteExcelData("Sheet1", row, 0, CampaignType);
		new CommonLib().WriteExcelData("Sheet1", row, cell, GetCount());
		logger.info("Fetch " + CampaignType + " Count");
		return Integer.parseInt(GetCount());

	}

	public void HoverMoreCampaign() throws Throwable {
		new CommonLib().MouseHover(GetMoreCampaign());
		new CommonLib().StandardWait(2000);
	}

	public void GetProgramLibrary(int row) throws Throwable {

		GetGlobalTreeSearch().sendKeys("Template");
		Thread.sleep(2000);
		try {
			driver.findElement(NoResult).isDisplayed();
			new CommonLib().WriteExcelData("Sheet1", row, 0, "Programe Library");
			new CommonLib().WriteExcelData("Sheet1", row, 1, "False");
			logger.info("No Template's are Present");

		} catch (Exception e) {
			new CommonLib().WriteExcelData("Sheet1", row, 0, "Programe Library");
			new CommonLib().WriteExcelData("Sheet1", row, 1, "True");
			screenshotUtility.TakeScreenshot(GetTemplate(), "Template");
			logger.info("Fetch Template's are Screenshot");

		}

	}

	public int GetMoreCampaignCount(String CampaignType, int row, int cell) throws Throwable {
		ClickCampaignInspector();
		HoverMoreCampaign();
		GetMoreCampaignOption(CampaignType);
		new CommonLib().waitForElementVisibleFlunt(AllCount);
		new CommonLib().WriteExcelData("Sheet1", row, 0, CampaignType);
		new CommonLib().WriteExcelData("Sheet1", row, cell, GetCount());
		int value = Integer.parseInt(GetCount());
		logger.info("Fetch " + CampaignType + " Count");
		return value;

	}

	public void GetEventCount(int row) throws Exception {
		new CommonLib().WaitForElementToLoad(driver, 60, GetFilter());
		GetFilter().click();
		new CommonLib().WaitForElementToLoad(driver, 60, GetEventFilter());
		GetEventFilter().click();
		try {
			boolean flag = driver.findElement(NoResult).isDisplayed();
			if (flag) {
				new CommonLib().WriteExcelData("Sheet1", row, 0, "Event Programs");
				new CommonLib().WriteExcelData("Sheet1", row, 1, "0");

			}

		} catch (Exception e) {
			List<WebElement> EventCount = driver.findElements(Event);
			new CommonLib().WriteExcelData("Sheet1", row, 0, "Event Programs");
			new CommonLib().WriteExcelData("Sheet1", row, 1, EventCount.size());
		}
		GetResetBtn().click();

	}

	public void GetNurtureCount(int row) throws Throwable {

		new CommonLib().WaitForElementToLoad(driver, 60, GetEngagementPrograms());
		GetEngagementPrograms().click();
		try {
			boolean flag = driver.findElement(NoResult).isDisplayed();
			if (flag) {
				new CommonLib().WriteExcelData("Sheet1", row, 0, "Nurture campaigns");
				new CommonLib().WriteExcelData("Sheet1", row, 1, "0");

			}

		} catch (Exception e) {
			List<WebElement> NurturCount = driver.findElements(Nuture);

			new CommonLib().WriteExcelData("Sheet1", row, 0, "Nurture campaigns");
			new CommonLib().WriteExcelData("Sheet1", row, 1, NurturCount.size());
		}
		GetResetBtn().click();
		GetFilter().click();

	}

	int cell = 2;

	int All_Triggered_Campaigns = 0;
	int Active_Triggered_Campaigns = 0;
	int Batch_Campaigns = 0;
	int All_Batch_Campaigns = 0;
	int AllCampaigns = 0;
	int Active_Campaigns = 0;

	public void AllWorkspaceCampaignCount() throws Throwable {
		List<WebElement> workSpace = driver.findElements(WorkSpace);

		new CommonLib().ClearExcelData("Sheet1", 7);
		new CommonLib().ClearExcelData("Sheet1", 8);
		new CommonLib().ClearExcelData("Sheet1", 9);
		new CommonLib().ClearExcelData("Sheet1", 10);
		new CommonLib().ClearExcelData("Sheet1", 11);
		new CommonLib().ClearExcelData("Sheet1", 12);
		new CommonLib().ClearExcelData("Sheet1", 13);
		new CommonLib().ClearExcelData("Sheet1", 26);

		for (WebElement value : workSpace) {
			logger.info("view " + value.getText() + " Workspace");
			value.click();
			switchFrame();
			All_Triggered_Campaigns += GetMoreCampaignCount("All Triggered Campaigns", 8, cell);
			Active_Triggered_Campaigns += GetMoreCampaignCount("Active Triggered Campaigns", 9, cell);
			Batch_Campaigns += GetMoreCampaignCount("Batch Campaigns - Repeating Schedule", 10, cell);
			All_Batch_Campaigns += GetMoreCampaignCount("All Batch Campaigns", 11, cell);
			AllCampaigns += GetCampaignCount("All Campaigns", 12, cell);
			Active_Campaigns += GetCampaignCount("Active Campaigns", 13, cell);
			driver.switchTo().defaultContent();
			new CommonLib().WriteExcelData("Sheet1", 7, 0, "Campaign Data");
			new CommonLib().WriteExcelData("Sheet1", 7, cell, value.getText());
			cell++;
			logger.info("Close " + value.getText() + " Workspace");

		}

		new CommonLib().WriteExcelData("Sheet1", 7, 1, "Total");
		new CommonLib().WriteExcelData("Sheet1", 26, 0, "Total WorkSpace");
		new CommonLib().WriteExcelData("Sheet1", 26, 1, workSpace.size());
		new CommonLib().WriteExcelData("Sheet1", 8, 1, All_Triggered_Campaigns);
		new CommonLib().WriteExcelData("Sheet1", 9, 1, Active_Triggered_Campaigns);
		new CommonLib().WriteExcelData("Sheet1", 10, 1, Batch_Campaigns);
		new CommonLib().WriteExcelData("Sheet1", 11, 1, All_Batch_Campaigns);
		new CommonLib().WriteExcelData("Sheet1", 12, 1, AllCampaigns);
		new CommonLib().WriteExcelData("Sheet1", 13, 1, Active_Campaigns);
		logger.info("Marketing Activite Page Task is done");

	}

	public void SpecificWorkspaceCampaignCount(int NoOfWorkSpace) throws Throwable {

		new CommonLib().ClearExcelData("Sheet1", 7);
		new CommonLib().ClearExcelData("Sheet1", 8);
		new CommonLib().ClearExcelData("Sheet1", 9);
		new CommonLib().ClearExcelData("Sheet1", 10);
		new CommonLib().ClearExcelData("Sheet1", 11);
		new CommonLib().ClearExcelData("Sheet1", 12);
		new CommonLib().ClearExcelData("Sheet1", 13);
		new CommonLib().ClearExcelData("Sheet1", 26);

		for (int i = 1; i <= NoOfWorkSpace; i++) {

			SelectWorkSpace(prop.getProperty("WorkSpace" + i));
			logger.info("view " + prop.getProperty("WorkSpace" + i) + " Workspace");

			switchFrame();
			All_Triggered_Campaigns += GetMoreCampaignCount("All Triggered Campaigns", 8, cell);
			Active_Triggered_Campaigns += GetMoreCampaignCount("Active Triggered Campaigns", 9, cell);
			Batch_Campaigns += GetMoreCampaignCount("Batch Campaigns - Repeating Schedule", 10, cell);
			All_Batch_Campaigns += GetMoreCampaignCount("All Batch Campaigns", 11, cell);
			AllCampaigns += GetCampaignCount("All Campaigns", 12, cell);
			Active_Campaigns += GetCampaignCount("Active Campaigns", 13, cell);
			driver.switchTo().defaultContent();
			new CommonLib().WriteExcelData("Sheet1", 7, 0, "Campaign Data");
			new CommonLib().WriteExcelData("Sheet1", 7, cell, prop.getProperty("WorkSpace" + i));
			cell++;
			logger.info("close " + prop.getProperty("WorkSpace" + i) + " Workspace");

		}

		new CommonLib().WriteExcelData("Sheet1", 7, 1, "Total");
		new CommonLib().WriteExcelData("Sheet1", 26, 0, "Total WorkSpace");
		new CommonLib().WriteExcelData("Sheet1", 26, 1, NoOfWorkSpace);
		new CommonLib().WriteExcelData("Sheet1", 8, 1, All_Triggered_Campaigns);
		new CommonLib().WriteExcelData("Sheet1", 9, 1, Active_Triggered_Campaigns);
		new CommonLib().WriteExcelData("Sheet1", 10, 1, Batch_Campaigns);
		new CommonLib().WriteExcelData("Sheet1", 11, 1, All_Batch_Campaigns);
		new CommonLib().WriteExcelData("Sheet1", 12, 1, AllCampaigns);
		new CommonLib().WriteExcelData("Sheet1", 13, 1, Active_Campaigns);

	}

}
