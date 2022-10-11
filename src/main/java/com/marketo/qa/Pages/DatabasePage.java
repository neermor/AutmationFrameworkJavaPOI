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

public class DatabasePage extends TestBase {
	MyMarketoPage homePage = new MyMarketoPage();
	private static Logger logger = LogManager.getLogger(TestBase.class);
	SoftAssert asrt = new SoftAssert();
	CommonLib Clib = new CommonLib();

	By TreeNode = By.xpath("//div[contains(@data-id,'treeNodeRow' )]");
	By Iframe = By.cssSelector("#mlm");
	By People = By.id("canvas__cp_ldbCanvasLeadList");
	By LeadsCount = By.cssSelector("[class='x-toolbar-right-row'] [class='xtb-text']");
	By Sag = By.xpath("//div[contains(@data-id,'treeNode_segmentation')]");
	By SagHar = By.cssSelector("#treeBodyAnchor > div > div > div:nth-child(6)");
	By WorkSpace = By.xpath("//div[contains(@data-id,'treeNode_workspace')]/..//div//span");

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
		if (words[0].equalsIgnoreCase("No")) {
			return 0;
		} else {
			return Integer.parseInt(words[2]);
		}

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

	public int GetLeadsCount(int row, int cell) throws Throwable {
		new CommonLib().StandardWait(2000);
		driver.switchTo().frame(GetIFrame());
		GetPeople().click();
		new CommonLib().WriteExcelData("Sheet1", row, 0, "Leads");
		new CommonLib().WriteExcelData("Sheet1", row, cell, GetCount());
		logger.info("Fetch Leads Count");

		return GetCount();
	}

	public int SegmentationsCount(int row, int cell) throws Throwable {
		try {
			boolean flag = driver.findElement(By.xpath(
					"//div[contains(@data-id,'treeNode_Label')]/span[text()='Segmentations']/../preceding-sibling::button[@data-id='treeNodeChevronIconButton']"))
					.isDisplayed();
			if (flag) {
				homePage.ExtendTreeNode("Segmentations");
				new CommonLib().WriteExcelData("Sheet1", row, 0, "Segmentations");
				new CommonLib().WriteExcelData("Sheet1", row, cell, GetSag().size());
				logger.info("Fetch Segmentations Count");
				new CommonLib().StandardWait(2000);
				screenshotUtility.TakeScreenshot(GetSagHar(), "Segmentations" + cell);
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

	int cell = 2;

	int Segmentations = 0;
	int Leads = 0;

	public void AllWorkspaceCollectLeadsCount() throws Throwable {

		List<WebElement> workSpace = driver.findElements(WorkSpace);
		new CommonLib().ClearExcelData("Sheet1", 14);
		new CommonLib().ClearExcelData("Sheet1", 15);
		new CommonLib().ClearExcelData("Sheet1", 16);

		for (WebElement value : workSpace) {
			try {
				GetExpandBtn(value.getText()).isDisplayed();
			} catch (Exception e) {
				Actions act = new Actions(driver);
				act.doubleClick(value).perform();
				logger.info("View " + value.getText() + " Workspace");

			}

			Segmentations += SegmentationsCount(15, cell);
			ExtendWorkshoptreenode("System Smart Lists", "All People");
			Leads += GetLeadsCount(16, cell);

			driver.switchTo().defaultContent();
			Actions act = new Actions(driver);
			act.doubleClick(value).perform();
			logger.info("Close " + value.getText() + " Workspace");

			new CommonLib().WriteExcelData("Sheet1", 14, 0, "Database Data");
			new CommonLib().WriteExcelData("Sheet1", 14, cell, value.getText());
			cell++;

		}
		new CommonLib().WriteExcelData("Sheet1", 14, 1, "Total");
		new CommonLib().WriteExcelData("Sheet1", 15, 1, Segmentations);
		new CommonLib().WriteExcelData("Sheet1", 16, 1, Leads);

	}

	boolean WorkspaceAvl = true;

	public void SpecificWorkspaceCollectLeadsCount(int NoOfWorkspace) throws Throwable {
		Clib.ClearExcelData("Sheet1", 14);
		Clib.ClearExcelData("Sheet1", 15);
		Clib.ClearExcelData("Sheet1", 16);

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

					Segmentations += SegmentationsCount(15, cell);
					ExtendWorkshoptreenode("System Smart Lists", "All People");
					Leads += GetLeadsCount(16, cell);

					driver.switchTo().defaultContent();
					act.doubleClick(workSpaceTree).perform();
					logger.info("Close " + ChooseWorkSpace(Workspace).getText() + " Workspace");

					new CommonLib().WriteExcelData("Sheet1", 14, 0, "Database Data");
					new CommonLib().WriteExcelData("Sheet1", 14, cell, workSpaceTree.getText());
					cell++;
				} catch (Exception ee) {
					logger.info("Oops!! " + Workspace + " Workspace is not available");
					asrt.assertTrue(WorkspaceAvl, Workspace + " Workspace is not available");

				}

			}
		}
		new CommonLib().WriteExcelData("Sheet1", 14, 1, "Total");
		new CommonLib().WriteExcelData("Sheet1", 15, 1, Segmentations);
		new CommonLib().WriteExcelData("Sheet1", 16, 1, Leads);
		asrt.assertAll();

	}
}
