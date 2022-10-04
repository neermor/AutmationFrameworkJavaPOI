package com.marketo.qa.Pages;

import java.util.List;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;

import com.marketo.qa.FileLib.CommonLib;
import com.marketo.qa.base.TestBase;
import com.marketo.qa.utility.screenshotUtility;

public class AnalyticsPage extends TestBase {

	MyMarketoPage homepage = new MyMarketoPage();
	private static Logger logger = LogManager.getLogger(TestBase.class);
	CommonLib Clib = new CommonLib();

	By Models = By.xpath("//div[contains(@data-id,'treeNode_revenuecyclemodel')]");
	By Rcm = By.cssSelector("#treeBodyAnchor > div > div > div:nth-child(4)");
	By WorkSpace = By.xpath("//div[contains(@data-id,'treeNode_workspace')]/..//div//span");

	public List<WebElement> GetModels() {
		return driver.findElements(Models);
	}

	public WebElement GetExpandBtn(String Name) {
		return driver.findElement(By.xpath("//div[contains(@data-id,'treeNode_workspace')]/..//div//span[text()='"
				+ Name
				+ "']/../preceding-sibling::button[contains(@data-id, 'treeNodeChevronIconButton')]//*[contains(@data-id, 'open')]"));
	}

	public WebElement GetRcm() {
		return driver.findElement(Rcm);
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

	public int ModelCount(int row, int cell) throws Throwable {
		homepage.ExtendTreeNode("Revenue Cycle Modeler");
		try {
			boolean flag = driver.findElement(By.xpath(
					"//div[contains(@data-id,'treeNode_Label')]/span[text()='Group Models']/../preceding-sibling::button[@data-id='treeNodeChevronIconButton']"))
					.isDisplayed();
			if (flag) {
				homepage.ExtendTreeNode("Group Models");
				new CommonLib().WriteExcelData("Sheet1", row, 0, "Models");
				new CommonLib().WriteExcelData("Sheet1", row, cell, GetModels().size());
				logger.info("Fetch Models Count");
				screenshotUtility.TakeScreenshot(GetRcm(), "Models" + cell);
				logger.info("Fetch Models Screenshot");
				return GetModels().size();
			}

		} catch (Exception e) {
			new CommonLib().WriteExcelData("Sheet1", row, 0, "Models");
			new CommonLib().WriteExcelData("Sheet1", row, cell, 0);
			logger.info("Models are Not Available");

		}
		return 0;

	}

	int cell = 2;

	int Model = 0;

	public void AllWorkspaceModelCount() throws Throwable {

		List<WebElement> workSpace = driver.findElements(WorkSpace);
		workSpace.size();
		new CommonLib().ClearExcelData("Sheet1", 17);
		new CommonLib().ClearExcelData("Sheet1", 18);

		for (WebElement value : workSpace) {
			try {
				if (GetExpandBtn(value.getText()).isDisplayed()) {
				}
			} catch (Exception e) {
				Actions act = new Actions(driver);
				act.doubleClick(value).perform();
				logger.info("View " + value.getText() + " Workspace");

			}

			Model += ModelCount(18, cell);

			driver.switchTo().defaultContent();
			Actions act = new Actions(driver);
			act.doubleClick(value).perform();
			logger.info("Close " + value.getText() + " Workspace");
			new CommonLib().WriteExcelData("Sheet1", 17, 0, "Program Data");
			new CommonLib().WriteExcelData("Sheet1", 17, cell, value.getText());
			cell++;

		}
		new CommonLib().WriteExcelData("Sheet1", 17, 1, "Total");
		new CommonLib().WriteExcelData("Sheet1", 18, 1, Model);

	}

	public void SpecificWorkspaceModelCount(int NoOfWorkspace) throws Throwable {

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

					Model += ModelCount(18, cell);

					driver.switchTo().defaultContent();
					act.doubleClick(workSpaceTree).perform();
					logger.info("Close " + ChooseWorkSpace(Workspace).getText() + " Workspace");

					new CommonLib().WriteExcelData("Sheet1", 17, 0, "Program Data");
					new CommonLib().WriteExcelData("Sheet1", 17, cell, workSpaceTree.getText());
					cell++;
				} catch (Exception ee) {
					driver.switchTo().defaultContent();
					logger.info("Oops!! " + Workspace + " Workspace is not available");
					e.printStackTrace();

				}
			}
		}
		new CommonLib().WriteExcelData("Sheet1", 17, 1, "Total");
		new CommonLib().WriteExcelData("Sheet1", 18, 1, Model);

	}

}
