package com.marketo.qa.Pages;

import java.util.Iterator;
import java.util.List;
import java.util.Set;

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

public class AnalyticsPage extends TestBase {

	MyMarketoPage homepage = new MyMarketoPage();
	private static Logger logger = LogManager.getLogger(TestBase.class);
	CommonLib Clib = new CommonLib();
	SoftAssert asrt = new SoftAssert();
	MarketingActivitePage mAP = new MarketingActivitePage();
	int i = 1;

	By Models = By.xpath("//div[contains(@data-id,'treeNode_revenuecyclemodel')]");
	By Rcm = By.cssSelector("#treeBodyAnchor > div > div > div:nth-child(4)");
	By WorkSpace = By.xpath("//div[contains(@data-id,'treeNode_workspace')]/..//div//span");
	By win = By.xpath("//*[@id=\"AppContainer\"]/div/div/div/div[2]/span");
	By ApprovedModels = By
			.xpath("//*[name()='use' and @*='#icon-meue-badge-green-check']/../../following-sibling::div/span");
	By PriviewBtn = By.xpath("//button[@class=' x-btn-text mkiPackageView']");
	By Toggle = By.xpath("//div[@class = 'x-tool x-tool-toggle x-tool-collapse-south']");
	By Modeler = By.xpath("//div[@class = 'WireIt-Layer']");

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

	public WebElement GetPBtn() {
		return driver.findElement(PriviewBtn);
	}

	public WebElement GetToggle() {
		return driver.findElement(Toggle);
	}

	public WebElement GetModeler() {
		return driver.findElement(Modeler);
	}

	String Parent_window = null;
	int cell = 3;
	int row = 33;

	public void FetchApprovedModelScreenshot(String WorkspaceName) throws Throwable {
		List<WebElement> ApprovedModel = driver.findElements(ApprovedModels);
		row++;
		System.out.println(ApprovedModel.size());
		Clib.WriteExcelData("Sheet1", row, 0, "Approved Models" + i);
		Clib.WriteExcelData("Sheet1", row, 2, WorkspaceName);
		Clib.WriteExcelData("Sheet1", row, 1, ApprovedModel.size());

		for (WebElement option : ApprovedModel) {

			String WindowTitle = option.getText();
			Clib.WriteExcelData("Sheet1", row, cell++, WindowTitle);

			option.click();
			mAP.switchFrame();
			Clib.StandardWait(2000);
			System.out.println(WindowTitle);
			GetPBtn().click();

			Set<String> s = driver.getWindowHandles();
			Iterator<String> I1 = s.iterator();

			while (I1.hasNext()) {
				Parent_window = I1.next();

				String child_window = I1.next();

				driver.switchTo().window(child_window);

				if (driver.getTitle().equalsIgnoreCase("Marketo | " + WindowTitle + " (Preview) • Analytics")) {

					Clib.StandardWait(6000);
					GetToggle().click();
					screenshotUtility.TakeScreenshotForModels(GetModeler(), WindowTitle);
					driver.close();
					driver.switchTo().window(Parent_window);

				}

			}

		}
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

	public int ModelCount(int row, int cell, String WorkSpace) throws Throwable {
		homepage.ExtendTreeNode("Revenue Cycle Modeler");
		try {
			boolean flag = driver.findElement(By.xpath(
					"//div[contains(@data-id,'treeNode_Label')]/span[text()='Group Models']/../preceding-sibling::button[@data-id='treeNodeChevronIconButton']"))
					.isDisplayed();
			if (flag) {
				homepage.ExtendTreeNode("Group Models");
				Clib.WriteExcelData("Sheet1", row, 0, "Models");
				Clib.WriteExcelData("Sheet1", row, cell, GetModels().size());
				Thread.sleep(1000);

				JavascriptExecutor js = (JavascriptExecutor) driver;
				WebElement element = driver.findElement(By.xpath("//div[@data-id='globalTreeDrawerExpanderContent']"));
				js.executeScript("arguments[0].setAttribute('style', 'width: 900px;')", element);
				screenshotUtility.TakeScreenshot(GetRcm(), WorkSpace);
				js.executeScript("arguments[0].setAttribute('style', 'width: 310px;')", element);
				try {
					FetchApprovedModelScreenshot(WorkSpace);
					i++;

				} catch (Exception e) {
					// TODO: handle exception
				}
				logger.info("Fetch Models Count");
				logger.info("Fetch Models Screenshot");
				return GetModels().size();
			}
		}

		catch (Exception e) {
			Clib.WriteExcelData("Sheet1", row, 0, "Models");
			Clib.WriteExcelData("Sheet1", row, cell, 0);
			logger.info("Models are Not Available");

		}
		return 0;

	}

	public void AllWorkspaceModelCount() throws Throwable {

		List<WebElement> workSpace = driver.findElements(WorkSpace);

		int cell = 2;
		int Model = 0;
		int ModelCount = 1;

		Clib.ClearExcelData("Sheet1", 17);
		Clib.ClearExcelData("Sheet1", 18);

		for (WebElement value : workSpace) {
			try {
				if (GetExpandBtn(value.getText()).isDisplayed()) {
				}
			} catch (Exception e) {
				Actions act = new Actions(driver);
				act.doubleClick(value).perform();
				logger.info("View " + value.getText() + " Workspace");

			}

			Model += ModelCount(18, ModelCount++, value.getText());

			driver.switchTo().defaultContent();
			Actions act = new Actions(driver);
			act.doubleClick(value).perform();
			logger.info("Close " + value.getText() + " Workspace");
			Clib.WriteExcelData("Sheet1", 17, 0, "Program Data");
			Clib.WriteExcelData("Sheet1", 17, cell, value.getText());
			cell++;

		}
		Clib.WriteExcelData("Sheet1", 17, 1, "Total");
		Clib.WriteExcelData("Sheet1", 18, 1, Model);

	}

	public void SpecificWorkspaceModelCount(int NoOfWorkspace) throws Throwable {

		int cell = 2;
		boolean WorkspaceAvl = true;
		int Model = 0;

		Clib.ClearExcelData("Sheet1", 17);
		Clib.ClearExcelData("Sheet1", 18);

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

					Model += ModelCount(18, cell, Workspace);

					driver.switchTo().defaultContent();
					act.doubleClick(workSpaceTree).perform();
					logger.info("Close " + ChooseWorkSpace(Workspace).getText() + " Workspace");

					Clib.WriteExcelData("Sheet1", 17, 0, "Program Data");
					Clib.WriteExcelData("Sheet1", 17, cell, workSpaceTree.getText());
					cell++;
				} catch (Exception ee) {
					driver.switchTo().defaultContent();
					logger.info("Oops!! " + Workspace + " Workspace is not available");
					e.printStackTrace();
					asrt.assertTrue(WorkspaceAvl, Workspace + " Workspace is not available");

				}
			}
		}
		Clib.WriteExcelData("Sheet1", 17, 1, "Total");
		Clib.WriteExcelData("Sheet1", 18, 1, Model);
		asrt.assertAll();
	}

}
