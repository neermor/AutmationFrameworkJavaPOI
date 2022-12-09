package com.marketo.qa.Pages;

import java.util.List;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.testng.asserts.SoftAssert;

import com.marketo.qa.FileLib.CommonLib;
import com.marketo.qa.base.TestBase;
import com.marketo.qa.utility.screenshotUtility;

public class DatabasePage extends TestBase {
	MyMarketoPage homePage = new MyMarketoPage();
	MarketingActivitePage mAP = new MarketingActivitePage();
	private static Logger logger = LogManager.getLogger(TestBase.class);
	SoftAssert asrt = new SoftAssert();
	CommonLib Clib = new CommonLib();
	int count = 0;
 
	By TreeNode = By.xpath("//div[contains(@data-id,'treeNodeRow' )]");
	By Iframe = By.cssSelector("#mlm");
	By People = By.id("canvas__cp_ldbCanvasLeadList");
	By LeadsCount = By.cssSelector("[class='x-toolbar-right-row'] [class='xtb-text']");
	By Sag = By.xpath("//div[contains(@data-id,'treeNode_segmentation')]/following-sibling::div");
	By SagHar = By.xpath("//span[text()=\"Segmentations\"]/../../../../..");
	By WorkSpace = By.xpath("//div[contains(@data-id,'treeNode_workspace')]/..//div//span");
	By Marketable = By.xpath("//*[@text='Marketable']/following-sibling:: *//*");

	public WebElement GetIFrame() {
		return driver.findElement(Iframe);
	}

	public WebElement GetExpandBtn(String Name) {
		return driver.findElement(By.xpath("//div[contains(@data-id,'treeNode_workspace')]/..//div//span[text()='"
				+ Name
				+ "']/../preceding-sibling::button[contains(@data-id, 'treeNodeChevronIconButton')]//*[contains(@data-id, 'open')]"));
	}

	public WebElement GetPeople() {
		return driver.findElement(People);
	}

	public WebElement GetMarketable() {
		return driver.findElement(Marketable);
	}

	public List<WebElement> GetSag() {
		return driver.findElements(Sag);
	}

	public WebElement GetSagHar() {
		return driver.findElement(SagHar);
	}

	WebElement workSpaceTree = null;

	public WebElement ChooseWorkSpace(String workSpaceName) {
		List<WebElement> workSpace = driver.findElements(WorkSpace);

		for (WebElement value : workSpace) {

			if (value.getText().equalsIgnoreCase(workSpaceName)) {
				workSpaceTree = value;
				break;
			}
		}
		return workSpaceTree;
	}

	public int GetCount() throws Throwable {

		Thread.sleep(4000);
		String countString = driver.findElement(LeadsCount).getText();
		String[] words = countString.split("\\s");
		if (words[0].startsWith("No")) {
			return 0;
		} else {
			if (mAP.GetPage().getText().equalsIgnoreCase("Of 1")) {
				count = Integer.parseInt(words[0]);

			} else {
				count = Integer.parseInt(words[2]);
			}
		}
		return count;

	}

	public void SelectTreeNode(String TreeNodeName) throws Throwable {
		Thread.sleep(10000);
		List<WebElement> HomeTiles = driver.findElements(TreeNode);
		for (WebElement option : HomeTiles) {

			if (option.getText().startsWith(TreeNodeName)) {
				option.click();
			}
		}
	}

	public void ExtendWorkshoptreenode(String Name, String ListName) throws Throwable {
		new CommonLib().StandardWait(2000);
		driver.findElement(By.xpath("//div[contains(@data-id,'treeNode_Label')]/span[text()=" + "'" + Name + "'"
				+ "]/../preceding-sibling::button[@data-id='treeNodeChevronIconButton']")).click();
		logger.info("View " + Name);
		SelectTreeNode(ListName);
		logger.info("Select " + ListName);
		new CommonLib().StandardWait(2000);
	}

	public int GetLeadsCount(int row, int cell, String option) throws Throwable {
		new CommonLib().StandardWait(2000);
		driver.switchTo().frame(GetIFrame());
		GetPeople().click();
		new CommonLib().StandardWait(4000);
		int value = GetCount();
		new CommonLib().WriteExcelData("Sheet1", row, 0, option);
		new CommonLib().WriteExcelData("Sheet1", row, cell, value);
		logger.info("Fetch Leads Count");
		driver.switchTo().defaultContent();

		return value;
	}

	public int SegmentationsCount(int row, int cell) throws Throwable {
		try {
			boolean flag = driver.findElement(By.xpath(
					"//div[contains(@data-id,'treeNode_Label')]/span[text()='Segmentations']/../preceding-sibling::button[@data-id='treeNodeChevronIconButton']"))
					.isDisplayed();
			if (flag) {
				homePage.ExtendTreeNode("Segmentations");
				new CommonLib().StandardWait(4000);
				new CommonLib().WriteExcelData("Sheet1", row, 0, "Segmentations");
				new CommonLib().WriteExcelData("Sheet1", row, cell, GetSag().size());
				logger.info("Fetch Segmentations Count");
				JavascriptExecutor js = (JavascriptExecutor) driver;
				new CommonLib().StandardWait(2000);
				WebElement element = driver.findElement(By.xpath("//div[@data-id='globalTreeDrawerExpanderContent']"));
				js.executeScript("arguments[0].setAttribute('style', 'width: 900px;')", element);
				screenshotUtility.TakeScreenshot(GetSagHar(), "Segmentations" + cell);
				js.executeScript("arguments[0].setAttribute('style', 'width: 310px;')", element);

				logger.info("Fetch Segmentations Screenshot");

				return GetSag().size();
			}
		} catch (Exception e) {
			new CommonLib().WriteExcelData("Sheet1", row, 0, "Segmentations");
			new CommonLib().WriteExcelData("Sheet1", row, cell, 0);
			logger.info("Segmentations are not available. .");

		}
		return 0;

	}

	public void AllWorkspaceCollectLeadsCount() throws Throwable {

		List<WebElement> workSpace = driver.findElements(WorkSpace);
		int cell = 2;
		int Segmentations = 0;
		int Leads = 0;
		int Marketable_Leads = 0;
		int Unsubscribed_Leads = 0;
		int Possible_Duplicates = 0;
		int Marketing_Suspended = 0;

		new CommonLib().ClearExcelData("Sheet1", 15);
		new CommonLib().ClearExcelData("Sheet1", 16);
		new CommonLib().ClearExcelData("Sheet1", 17);
		new CommonLib().ClearExcelData("Sheet1", 18);
		new CommonLib().ClearExcelData("Sheet1", 19);
		new CommonLib().ClearExcelData("Sheet1", 20);
		new CommonLib().ClearExcelData("Sheet1", 21);

		for (WebElement value : workSpace) {
			try {
				GetExpandBtn(value.getText()).isDisplayed();
			} catch (Exception e) {
				Actions act = new Actions(driver);
				act.doubleClick(value).perform();
				logger.info("View " + value.getText() + " Workspace");

			}

			Segmentations += SegmentationsCount(16, cell);
			driver.switchTo().frame(GetIFrame());
			Marketable_Leads = Integer.parseInt(GetMarketable().getText());
			new CommonLib().WriteExcelData("Sheet1", 21, 0, "Marketable Leads");
			new CommonLib().WriteExcelData("Sheet1", 21, cell, GetSag().size());
			driver.switchTo().defaultContent();
			ExtendWorkshoptreenode("System Smart Lists", "All People");
			Leads += GetLeadsCount(17, cell, "All People");
			Marketable_Leads = Integer.parseInt(GetMarketable().getText());
			new CommonLib().WriteExcelData("Sheet1", 21, 0, "Marketable Leads");
			new CommonLib().WriteExcelData("Sheet1", 21, cell, GetSag().size());
			SelectTreeNode("Unsubscribed People");
			Unsubscribed_Leads += GetLeadsCount(18, cell, "Unsubscribed People");
			SelectTreeNode("Possible Duplicates");
			Possible_Duplicates += GetLeadsCount(19, cell, "Possible Duplicates");
			SelectTreeNode("Marketing Suspended");
			Marketing_Suspended += GetLeadsCount(20, cell, "Marketing Suspended");

			driver.switchTo().defaultContent();
			Actions act = new Actions(driver);
			act.doubleClick(value).perform();
			logger.info("Close " + value.getText() + " Workspace");

			new CommonLib().WriteExcelData("Sheet1", 15, 0, "Database Data");
			new CommonLib().WriteExcelData("Sheet1", 15, cell, value.getText());
			cell++;

		}
		new CommonLib().WriteExcelData("Sheet1", 15, 1, "Total");
		new CommonLib().WriteExcelData("Sheet1", 16, 1, Segmentations);
		new CommonLib().WriteExcelData("Sheet1", 17, 1, Leads);
		new CommonLib().WriteExcelData("Sheet1", 18, 1, Unsubscribed_Leads);
		new CommonLib().WriteExcelData("Sheet1", 19, 1, Possible_Duplicates);
		new CommonLib().WriteExcelData("Sheet1", 20, 1, Marketing_Suspended);
		new CommonLib().WriteExcelData("Sheet1", 21, 1, Marketable_Leads);

	}

	public void SpecificWorkspaceCollectLeadsCount(int NoOfWorkspace) throws Throwable {

		new CommonLib().ClearExcelData("Sheet1", 15);
		new CommonLib().ClearExcelData("Sheet1", 16);
		new CommonLib().ClearExcelData("Sheet1", 17);
		new CommonLib().ClearExcelData("Sheet1", 18);
		new CommonLib().ClearExcelData("Sheet1", 19);
		new CommonLib().ClearExcelData("Sheet1", 20);
		new CommonLib().ClearExcelData("Sheet1", 21);

		boolean WorkspaceAvl = true;
		int cell = 2;
		int Segmentations = 0;
		int Leads = 0;
		int Unsubscribed_Leads = 0;
		int Possible_Duplicates = 0;
		int Marketing_Suspended = 0;
		int Marketable_Leads = 0;

		new DesignStudioPage().CloseDefaultTreeView();
		for (int i = 1; i <= NoOfWorkspace; i++) {
			String Workspace = prop.getProperty("WorkSpace" + i);
			try {
				GetExpandBtn(Workspace).isDisplayed();

			} catch (Exception e) {
				Actions act = new Actions(driver);
				try {
					act.doubleClick(ChooseWorkSpace(Workspace)).perform();
					logger.info("View " + ChooseWorkSpace(Workspace).getText() + " Workspace");

					Segmentations += SegmentationsCount(16, cell);
					driver.switchTo().frame(GetIFrame());
					String value = GetMarketable().getText();
					value = value.replace(",", "");
					Marketable_Leads = Integer.parseInt(value);
					new CommonLib().WriteExcelData("Sheet1", 21, 0, "Marketable Leads");
					new CommonLib().WriteExcelData("Sheet1", 21, cell, GetSag().size());
					driver.switchTo().defaultContent();

					ExtendWorkshoptreenode("System Smart Lists", "All People");

					Leads += GetLeadsCount(17, cell, "All People");
					SelectTreeNode("Unsubscribed People");
					Unsubscribed_Leads += GetLeadsCount(18, cell, "Unsubscribed People");
					SelectTreeNode("Possible Duplicates");
					Possible_Duplicates += GetLeadsCount(19, cell, "Possible Duplicates");
					SelectTreeNode("Marketing Suspended");
					Marketing_Suspended += GetLeadsCount(20, cell, "Marketing Suspended");

					driver.switchTo().defaultContent();
					act.doubleClick(workSpaceTree).perform();
					logger.info("Close " + ChooseWorkSpace(Workspace).getText() + " Workspace");

					new CommonLib().WriteExcelData("Sheet1", 15, 0, "Database Data");
					new CommonLib().WriteExcelData("Sheet1", 15, cell, workSpaceTree.getText());
					cell++;
				} catch (Exception ee) {
					logger.info("Oops!! " + Workspace + " Workspace is not available");
					asrt.assertTrue(WorkspaceAvl, Workspace + " Workspace is not available");

				}

			}
		}
		new CommonLib().WriteExcelData("Sheet1", 15, 1, "Total");
		new CommonLib().WriteExcelData("Sheet1", 16, 1, Segmentations);
		new CommonLib().WriteExcelData("Sheet1", 17, 1, Leads);
		new CommonLib().WriteExcelData("Sheet1", 18, 1, Unsubscribed_Leads);
		new CommonLib().WriteExcelData("Sheet1", 19, 1, Possible_Duplicates);
		new CommonLib().WriteExcelData("Sheet1", 20, 1, Marketing_Suspended);
		new CommonLib().WriteExcelData("Sheet1", 21, 1, Marketable_Leads);

		asrt.assertAll();

	}
}
