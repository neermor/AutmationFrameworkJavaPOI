package com.marketo.qa.Pages;

import java.util.List;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;

import com.marketo.qa.FileLib.CommonLib;
import com.marketo.qa.base.TestBase;

public class DesignStudioPage extends TestBase {
	MyMarketoPage homepage = new MyMarketoPage();
	MarketingActivitePage mAP = new MarketingActivitePage();
	CommonLib Clib = new CommonLib();
	private static Logger logger = LogManager.getLogger(TestBase.class);

	By TreeNode = By.xpath("//div[contains(@data-id,'treeNodeRow' )]");
	By AllCount = By.cssSelector("[data-id='dataGridFooter_pageInfo']");
	By UploadCount = By.cssSelector(".x-toolbar-right-row  .x-toolbar-cell .xtb-text");
	By Iframe = By.cssSelector("#mlm");
	By WorkSpace = By.xpath("//div[contains(@data-id,'treeNode_workspace')]/..//div//span");
	By DefaultWP = By.xpath("//div[@data-id='treeNode_Label']//span[text()='Default']");
	By CountNxtPage = By.xpath("//button[@data-id='dataGridFooter_nextPage']");
	By Refresh = By.xpath("//button[@data-id='dataGridFooter_refresh']");

	public WebElement GetIFrame() {
		return driver.findElement(Iframe);
	}

	public WebElement GetDefaultWP() {
		return driver.findElement(DefaultWP);
	}

	public WebElement GetRefresh() {
		return driver.findElement(Refresh);
	}

	public WebElement GetNextPage() {
		return driver.findElement(CountNxtPage);
	}

	public WebElement GetExpandBtn(String Name) {
		return driver.findElement(By.xpath("//div[contains(@data-id,'treeNode_workspace')]/..//div//span[text()='"
				+ Name
				+ "']/../preceding-sibling::button[contains(@data-id, 'treeNodeChevronIconButton')]//*[contains(@data-id, 'open')]"));
	}

	boolean flag = true;

	public int GetCount() throws Throwable {
		int count = 0;

		try {
			while (flag) {
				GetNextPage().click();
				Clib.WaitForElementToLoad(driver, 40, GetNextPage());
				flag = GetNextPage().isEnabled();
			}
		} catch (Exception e) {
			Clib.StandardWait(2000);
			GetRefresh().click();
			Clib.StandardWait(2000);
			String countString = driver.findElement(AllCount).getText();
			String[] words = countString.split("\\s");
			if (words[0].equalsIgnoreCase("0")) {
				return 0;
			}
			count = Integer.parseInt(words[2]);
			System.out.println(count);

		}
		flag = true;
		return count;
	}

	public int GetUploadDataCount() throws Throwable {
		Clib.StandardWait(4000);
		String countString = driver.findElement(UploadCount).getText();
		String[] words = countString.split("\\s");
		if (countString.equalsIgnoreCase("No images or files exist")) {
			return 0;
		} else {
			return Integer.parseInt(words[1]);
		}
	}

	public String GetSnippetsCount() throws Throwable {
		Clib.StandardWait(10000);
		String countString = driver.findElement(UploadCount).getText();
		String[] words = countString.split("\\s");
		if (countString.equalsIgnoreCase("No snippets exist")) {
			return "0";
		} else {
			if (mAP.GetPage().getText().equalsIgnoreCase("Of 1")) {
				return words[0];
			} else {
				return words[2];
			}
		}

	}

	public int FetchTreeNodeCount(String value, int row, int cell) throws Throwable {
		homepage.SelectTreeNode(value);
		logger.info("Select " + value);
		Clib.StandardWait(3000);
		int count = GetCount();
		Clib.WriteExcelData("Sheet1", row, 0, value);
		Clib.WriteExcelData("Sheet1", row, cell, count);
		logger.info("Fetch " + value + " Count");
		return count;
	}

	public int FetchUploadCount(String value, int row, int cell) throws Throwable {
		driver.switchTo().defaultContent();
		homepage.SelectTreeNode(value);
		driver.switchTo().frame(GetIFrame());
		Clib.StandardWait(2000);
		Clib.WriteExcelData("Sheet1", row, 0, value);
		Clib.WriteExcelData("Sheet1", row, cell, GetUploadDataCount());
		return GetUploadDataCount();

	}

	public int FetchSnippetsCount(String value, int row, int cell) throws Throwable {
		driver.switchTo().defaultContent();
		homepage.SelectTreeNode(value);
		logger.info("Select " + value);
		driver.switchTo().frame(GetIFrame());
		Clib.StandardWait(2000);
		Clib.WriteExcelData("Sheet1", row, 0, value);
		Clib.WriteExcelData("Sheet1", row, cell, GetSnippetsCount());
		logger.info("Fetch " + value + " Count");
		int SnippetsCount = Integer.parseInt(GetSnippetsCount());
		return SnippetsCount;

	}

	WebElement workSpaceTree = null;

	public WebElement ChooseWorkSpace(String workSpaceName) {
		List<WebElement> workSpace = driver.findElements(WorkSpace);

		for (WebElement value : workSpace) {
			workSpaceTree = null;
			if (value.getText().equalsIgnoreCase(workSpaceName)) {
				workSpaceTree = value;
				break;
			}
		}
		if (workSpaceTree == null) {
			logger.info("Opps!! WorkSpace is Not Available");
		}
		return workSpaceTree;

	}

	int cell = 2;

	int AllEmails = 0;
	int AllForms = 0;
	int AllLandingPages = 0;
	int AllImages_and_Files = 0;
	int AllSnippets = 0;

	public void AllWorkspaceRequiredCount() throws Throwable {

		List<WebElement> workSpace = driver.findElements(WorkSpace);

		Clib.ClearExcelData("Sheet1", 1);
		Clib.ClearExcelData("Sheet1", 2);
		Clib.ClearExcelData("Sheet1", 3);
		Clib.ClearExcelData("Sheet1", 4);
		Clib.ClearExcelData("Sheet1", 5);
		Clib.ClearExcelData("Sheet1", 6);

		for (WebElement value : workSpace) {
			try {
				GetExpandBtn(value.getText()).isDisplayed();

			} catch (Exception e) {
				Actions act = new Actions(driver);
				act.doubleClick(value).perform();
				logger.info("View " + value.getText() + " Workspace");

			}

			AllEmails += FetchTreeNodeCount("Emails", 2, cell);
			AllForms += FetchTreeNodeCount("Forms", 3, cell);
			AllLandingPages += FetchTreeNodeCount("Landing Pages", 4, cell);
			AllImages_and_Files += FetchTreeNodeCount("Images and Files", 5, cell);
			AllSnippets += FetchSnippetsCount("Snippets", 6, cell);

			driver.switchTo().defaultContent();

			Actions act = new Actions(driver);
			act.doubleClick(value).perform();
			logger.info("Close " + value.getText() + " Workspace");

			Clib.WriteExcelData("Sheet1", 1, 0, "Asset Data");
			Clib.WriteExcelData("Sheet1", 1, cell, value.getText());
			cell++;

		}

		Clib.WriteExcelData("Sheet1", 1, 1, "Total");
		Clib.WriteExcelData("Sheet1", 2, 1, AllEmails);
		Clib.WriteExcelData("Sheet1", 3, 1, AllForms);
		Clib.WriteExcelData("Sheet1", 4, 1, AllLandingPages);
		Clib.WriteExcelData("Sheet1", 5, 1, AllImages_and_Files);
		Clib.WriteExcelData("Sheet1", 6, 1, AllSnippets);

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

	public void SpecificWorkspaceRequiredCount(int NoOfWorkspace) throws Throwable {

		Clib.ClearExcelData("Sheet1", 1);
		Clib.ClearExcelData("Sheet1", 2);
		Clib.ClearExcelData("Sheet1", 3);
		Clib.ClearExcelData("Sheet1", 4);
		Clib.ClearExcelData("Sheet1", 5);
		Clib.ClearExcelData("Sheet1", 6);

		CloseDefaultTreeView();
		for (int i = 1; i <= NoOfWorkspace; i++) {
			String Workspace = prop.getProperty("WorkSpace" + i);
			try {
				GetExpandBtn(Workspace).isDisplayed();

			} catch (Exception e) {
				Actions act = new Actions(driver);
				try {
					act.doubleClick(ChooseWorkSpace(Workspace)).perform();
					logger.info("view " + ChooseWorkSpace(Workspace).getText() + " Workspace");

					AllForms += FetchTreeNodeCount("Forms", 3, cell);
					AllLandingPages += FetchTreeNodeCount("Landing Pages", 4, cell);
					AllEmails += FetchTreeNodeCount("Emails", 2, cell);
					AllImages_and_Files += FetchTreeNodeCount("Images and Files", 5, cell);
					AllSnippets += FetchSnippetsCount("Snippets", 6, cell);

					driver.switchTo().defaultContent();

					act.doubleClick(ChooseWorkSpace(Workspace)).perform();
					logger.info("Close " + ChooseWorkSpace(Workspace).getText() + " Workspace");

					Clib.WriteExcelData("Sheet1", 1, 0, "Asset Data");
					Clib.WriteExcelData("Sheet1", 1, cell, ChooseWorkSpace(Workspace).getText());
					cell++;
				} catch (Exception ee) {
					logger.info("Opps!! " + Workspace + " Workspace is not available");
				}
			}
		}

		Clib.WriteExcelData("Sheet1", 1, 1, "Total");
		Clib.WriteExcelData("Sheet1", 2, 1, AllEmails);
		Clib.WriteExcelData("Sheet1", 3, 1, AllForms);
		Clib.WriteExcelData("Sheet1", 4, 1, AllLandingPages);
		Clib.WriteExcelData("Sheet1", 5, 1, AllImages_and_Files);
		Clib.WriteExcelData("Sheet1", 6, 1, AllSnippets);

	}

}
