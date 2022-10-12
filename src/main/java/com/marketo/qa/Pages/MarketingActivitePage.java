package com.marketo.qa.Pages;

import java.util.List;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.testng.asserts.SoftAssert;

import com.marketo.qa.FileLib.CommonLib;
import com.marketo.qa.base.TestBase;
import com.marketo.qa.utility.screenshotUtility;

public class MarketingActivitePage extends TestBase {
	private static Logger logger = LogManager.getLogger(TestBase.class);
	CommonLib Clib = new CommonLib();
	SoftAssert asrt = new SoftAssert();

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
	By DefaultWP = By.xpath("//div[@data-id='treeNode_Label']//span[text()='Default']");

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

	public WebElement GetDefaultWP() {
		return driver.findElement(DefaultWP);
	}

	public WebElement GetMoreCampaign() {
		return driver.findElement(MoreCampaigns);
	}

	public void switchFrame() throws Throwable {
		driver.switchTo().frame(GetIFrame());

	}

	public WebElement GetExpandBtn(String Name) {
		return driver.findElement(By.xpath("//div[contains(@data-id,'treeNode_workspace')]/..//div//span[text()='"
				+ Name
				+ "']/../preceding-sibling::button[contains(@data-id, 'treeNodeChevronIconButton')]//*[contains(@data-id, 'open')]"));
	}

	public void ClickCampaignInspector() throws Throwable {
		Clib.StandardWait(2000);
		GetCampaignInspector().click();
		GetCampaignDD().click();

	}

	public boolean CheckDataAvailable(String Name) {
		boolean Value = false;
		try {
			Value = driver
					.findElement(By.xpath("//div[contains(@data-id,'treeNode_workspace')]/..//div//span[text()='" + Name
							+ "']/../preceding-sibling::button[contains(@data-id, 'treeNodeChevronIconButton')]"))
					.isDisplayed();

		} catch (Exception e) {
			logger.info(Name + " Data is not available");
		}
		return Value;
	}

	public WebElement SelectTreeNode(String TreeNodeName) throws Throwable {
		Clib.StandardWait(2000);
		List<WebElement> HomeTiles = driver.findElements(TreeNode);
		WebElement wb = null;
		for (WebElement option : HomeTiles) {

			if (option.getText().equalsIgnoreCase(TreeNodeName)) {
				wb = option;
			}
		}
		return wb;
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
		Clib.WriteExcelData("Sheet1", row, 0, CampaignType);
		Clib.WriteExcelData("Sheet1", row, cell, GetCount());
		logger.info("Fetch " + CampaignType + " Count");
		return Integer.parseInt(GetCount());

	}

	public void HoverMoreCampaign() throws Throwable {
		Clib.MouseHover(GetMoreCampaign());
		Clib.StandardWait(2000);
	}

	public void GetProgramLibrary(int row) throws Throwable {

		GetGlobalTreeSearch().sendKeys("Template");
		Thread.sleep(2000);
		try {
			driver.findElement(NoResult).isDisplayed();
			Clib.WriteExcelData("Sheet1", row, 0, "Programe Library");
			Clib.WriteExcelData("Sheet1", row, 1, "False");
			logger.info("No Template's are Present");

		} catch (Exception e) {
			Clib.WriteExcelData("Sheet1", row, 0, "Programe Library");
			Clib.WriteExcelData("Sheet1", row, 1, "True");
			screenshotUtility.TakeScreenshot(GetTemplate(), "Template");
			logger.info("Fetch Template's Screenshot");

		}

	}

	public int GetMoreCampaignCount(String CampaignType, int row, int cell) throws Throwable {
		ClickCampaignInspector();
		HoverMoreCampaign();
		GetMoreCampaignOption(CampaignType);
		Clib.waitForElementVisibleFlunt(AllCount);
		Clib.WriteExcelData("Sheet1", row, 0, CampaignType);
		Clib.WriteExcelData("Sheet1", row, cell, GetCount());
		int value = Integer.parseInt(GetCount());
		logger.info("Fetch " + CampaignType + " Count");
		return value;

	}

	public void GetEventCount(int row) throws Exception {
		Clib.WaitForElementToLoad(driver, 60, GetFilter());
		GetFilter().click();
		Clib.WaitForElementToLoad(driver, 60, GetEventFilter());
		GetEventFilter().click();
		try {
			boolean flag = driver.findElement(NoResult).isDisplayed();
			if (flag) {
				Clib.WriteExcelData("Sheet1", row, 0, "Event Programs");
				Clib.WriteExcelData("Sheet1", row, 1, "0");

			}

		} catch (Exception e) {
			List<WebElement> EventCount = driver.findElements(Event);
			Clib.WriteExcelData("Sheet1", row, 0, "Event Programs");
			Clib.WriteExcelData("Sheet1", row, 1, EventCount.size());
		}
		GetResetBtn().click();

	}

	public void GetNurtureCount(int row) throws Throwable {

		Clib.WaitForElementToLoad(driver, 60, GetEngagementPrograms());
		GetEngagementPrograms().click();
		try {
			boolean flag = driver.findElement(NoResult).isDisplayed();
			if (flag) {
				Clib.WriteExcelData("Sheet1", row, 0, "Nurture campaigns");
				Clib.WriteExcelData("Sheet1", row, 1, "0");

			}

		} catch (Exception e) {
			List<WebElement> NurturCount = driver.findElements(Nuture);

			Clib.WriteExcelData("Sheet1", row, 0, "Nurture campaigns");
			Clib.WriteExcelData("Sheet1", row, 1, NurturCount.size());
		}
		GetResetBtn().click();
		GetFilter().click();

	}

	public void CloseDefaultTreeView() throws Throwable {
		Clib.StandardWait(4000);
		try {
			GetExpandBtn("Default").isDisplayed();
			Actions act = new Actions(driver);
			act.doubleClick(GetDefaultWP()).perform();
		} catch (Exception e) {
			// TODO: handle exception
		}

	}

	public void AllWorkspaceCampaignCount() throws Throwable {
		List<WebElement> workSpace = driver.findElements(WorkSpace);

		int cell = 2;
		boolean WorkspaceAvl = true;
		int All_Triggered_Campaigns = 0;
		int Active_Triggered_Campaigns = 0;
		int Batch_Campaigns = 0;
		int All_Batch_Campaigns = 0;
		int AllCampaigns = 0;
		int Active_Campaigns = 0;

		Clib.ClearExcelData("Sheet1", 7);
		Clib.ClearExcelData("Sheet1", 8);
		Clib.ClearExcelData("Sheet1", 9);
		Clib.ClearExcelData("Sheet1", 10);
		Clib.ClearExcelData("Sheet1", 11);
		Clib.ClearExcelData("Sheet1", 12);
		Clib.ClearExcelData("Sheet1", 13);
		Clib.ClearExcelData("Sheet1", 26);

		for (WebElement value : workSpace) {
			logger.info("view " + value.getText() + " Workspace");
			System.out.println(CheckDataAvailable(value.getText()));
			if (CheckDataAvailable(value.getText())) {
				try {
					value.click();
					switchFrame();
					All_Triggered_Campaigns += GetMoreCampaignCount("All Triggered Campaigns", 8, cell);
					Active_Triggered_Campaigns += GetMoreCampaignCount("Active Triggered Campaigns", 9, cell);
					Batch_Campaigns += GetMoreCampaignCount("Batch Campaigns - Repeating Schedule", 10, cell);
					All_Batch_Campaigns += GetMoreCampaignCount("All Batch Campaigns", 11, cell);
					AllCampaigns += GetCampaignCount("All Campaigns", 12, cell);
					Active_Campaigns += GetCampaignCount("Active Campaigns", 13, cell);
					driver.switchTo().defaultContent();
					Clib.WriteExcelData("Sheet1", 7, 0, "Campaign Data");
					Clib.WriteExcelData("Sheet1", 7, cell, value.getText());
					cell++;
					logger.info("Close " + value.getText() + " Workspace");

				} catch (Exception e) {
					driver.switchTo().defaultContent();
					logger.info("Oops!! " + value.getText() + " Workspace is not available");
					e.printStackTrace();
					WorkspaceAvl = false;
					asrt.assertTrue(WorkspaceAvl, value.getText() + " Workspace is not available");
				}
			}
		}

		Clib.WriteExcelData("Sheet1", 7, 1, "Total");
		Clib.WriteExcelData("Sheet1", 26, 0, "Total WorkSpace");
		Clib.WriteExcelData("Sheet1", 26, 1, workSpace.size());
		Clib.WriteExcelData("Sheet1", 8, 1, All_Triggered_Campaigns);
		Clib.WriteExcelData("Sheet1", 9, 1, Active_Triggered_Campaigns);
		Clib.WriteExcelData("Sheet1", 10, 1, Batch_Campaigns);
		Clib.WriteExcelData("Sheet1", 11, 1, All_Batch_Campaigns);
		Clib.WriteExcelData("Sheet1", 12, 1, AllCampaigns);
		Clib.WriteExcelData("Sheet1", 13, 1, Active_Campaigns);
		logger.info("Marketing Activite Page Task is done");

	}

	public void SpecificWorkspaceCampaignCount(int NoOfWorkSpace) throws Throwable {

		int cell = 2;
		boolean WorkspaceAvl = true;
		int All_Triggered_Campaigns = 0;
		int Active_Triggered_Campaigns = 0;
		int Batch_Campaigns = 0;
		int All_Batch_Campaigns = 0;
		int AllCampaigns = 0;
		int Active_Campaigns = 0;

		Clib.ClearExcelData("Sheet1", 7);
		Clib.ClearExcelData("Sheet1", 8);
		Clib.ClearExcelData("Sheet1", 9);
		Clib.ClearExcelData("Sheet1", 10);
		Clib.ClearExcelData("Sheet1", 11);
		Clib.ClearExcelData("Sheet1", 12);
		Clib.ClearExcelData("Sheet1", 13);
		Clib.ClearExcelData("Sheet1", 26);

		driver.switchTo().defaultContent();

		CloseDefaultTreeView();

		for (int i = 1; i <= NoOfWorkSpace; i++) {
			String Workspace = prop.getProperty("WorkSpace" + i);
			System.out.println(CheckDataAvailable(Workspace));

			if (CheckDataAvailable(Workspace)) {
				try {
					CheckDataAvailable(Workspace);
					WebElement wSpace = SelectTreeNode(Workspace);
					wSpace.click();
					logger.info("view " + Workspace + " Workspace");
					switchFrame();
					All_Triggered_Campaigns += GetMoreCampaignCount("All Triggered Campaigns", 8, cell);
					Active_Triggered_Campaigns += GetMoreCampaignCount("Active Triggered Campaigns", 9, cell);
					Batch_Campaigns += GetMoreCampaignCount("Batch Campaigns - Repeating Schedule", 10, cell);
					All_Batch_Campaigns += GetMoreCampaignCount("All Batch Campaigns", 11, cell);
					AllCampaigns += GetCampaignCount("All Campaigns", 12, cell);
					Active_Campaigns += GetCampaignCount("Active Campaigns", 13, cell);
					driver.switchTo().defaultContent();
					Clib.WriteExcelData("Sheet1", 7, 0, "Campaign Data");
					Clib.WriteExcelData("Sheet1", 7, cell, prop.getProperty("WorkSpace" + i));
					cell++;
					logger.info("close " + Workspace + " Workspace");
				} catch (Exception e) {
					driver.switchTo().defaultContent();
					logger.info("Oops!! " + Workspace + " Workspace is not available");
					e.printStackTrace();
					WorkspaceAvl = false;
					asrt.assertTrue(WorkspaceAvl, Workspace + " Workspace is not available");
				}
			}

		}

		Clib.WriteExcelData("Sheet1", 7, 1, "Total");
		Clib.WriteExcelData("Sheet1", 26, 0, "Total WorkSpace");
		Clib.WriteExcelData("Sheet1", 26, 1, NoOfWorkSpace);
		Clib.WriteExcelData("Sheet1", 8, 1, All_Triggered_Campaigns);
		Clib.WriteExcelData("Sheet1", 9, 1, Active_Triggered_Campaigns);
		Clib.WriteExcelData("Sheet1", 10, 1, Batch_Campaigns);
		Clib.WriteExcelData("Sheet1", 11, 1, All_Batch_Campaigns);
		Clib.WriteExcelData("Sheet1", 12, 1, AllCampaigns);
		Clib.WriteExcelData("Sheet1", 13, 1, Active_Campaigns);
		asrt.assertAll();

	}

}
