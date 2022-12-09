package com.marketo.qa.Pages;

import java.util.Iterator;
import java.util.List;
import java.util.Set;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.testng.Reporter;

import com.marketo.qa.FileLib.CommonLib;
import com.marketo.qa.base.TestBase;
import com.marketo.qa.utility.screenshotUtility;

public class WebPersonalization extends TestBase {
	MyMarketoPage homepage = new MyMarketoPage();
	private static Logger logger = LogManager.getLogger(TestBase.class);
	CommonLib Clib = new CommonLib();
	MarketingActivitePage mAP = new MarketingActivitePage();

	By TopIndustries = By.xpath("//h4[text()='Top Industries'] /../../../../..");
	By MktoBall = By.id("mktoBall");
	By Dashboard = By.xpath("//a[@class='navMenuIcon menuItemDashboard']");
	By WebCampaigns = By.xpath("//a[@class='navMenuIcon menuItemCtas']");
	By CampaignsList = By.id("campaignsGroupContainer");
	By WorkSpace = By.xpath("//*[@class='navMenu']");
	By TotalOrgCount = By.xpath("//div[@id='totalOrganizationCounter']");
	By WSDP = By.id("workspacesMenuLabel");
	By WebCampaignsNoResults = By.xpath("//div[@class='sort-list']//div[contains(text(),'Displaying')]");

	public WebElement GetTopIndustries() {
		return driver.findElement(TopIndustries);
	}

	public WebElement GetWebCampaignsNoResults() {
		return driver.findElement(WebCampaignsNoResults);
	}

	public WebElement GetWorkSpaceDropDown() {
		return driver.findElement(WSDP);
	}

	public WebElement GetTotalOrgCount() {
		return driver.findElement(TotalOrgCount);
	}

	public WebElement GetMktoBall() {
		return driver.findElement(MktoBall);
	}

	public WebElement GetDashboard() {
		return driver.findElement(Dashboard);
	}

	public WebElement GetWebCampaigns() {
		return driver.findElement(WebCampaigns);
	}

	public WebElement GetCampaignsList() {
		return driver.findElement(CampaignsList);
	}

	public WebElement GetDashboard(String Name) {
		return driver.findElement(By.xpath("//h4[text()='" + Name + "'] /../../../../../.."));
	};

	public WebElement GetTotalOrganizations_Top_Organizations() {
		return driver.findElement(By.xpath("//h4[text()='Total Organizations'] /../../../../../../.."));
	};

	public boolean VerifyAndFetchScreenshots(int row) throws Throwable {
		Clib.ClearExcelData("Sheet1", row);

		try {

			homepage.GetHometileUnderFrame("Web Personalization").click();
			Set<String> s = driver.getWindowHandles();
			Iterator<String> I1 = s.iterator();

			Parent_window = I1.next();
			String child_window = I1.next();
			driver.switchTo().window(child_window);
			if (driver.getTitle().equalsIgnoreCase("Marketo | Dashboard")) {
				Clib.WriteExcelData("Sheet1", row, 0, "Web Personalization");
				Clib.WriteExcelData("Sheet1", row, 1, "True");
				return true;
			} else {
				Clib.WriteExcelData("Sheet1", row, 0, "Web Personalization");
				Clib.WriteExcelData("Sheet1", row, 1, "False");
				logger.info("Web Personalization not available");
				return false;

			}
		}

		catch (Exception e) {
			Clib.WriteExcelData("Sheet1", row, 0, "Web Personalization");
			Clib.WriteExcelData("Sheet1", row, 1, "False");
			logger.info("Web Personalization not available");

			return false;
		}
	}

	String Parent_window = null;

	public void WebCampaignsScreenShot(int row, String WorkspaceCondition, int TotalOrgCountrow) throws Throwable {
		int size = 1;
		for (int i = 0; i < size; i++) {
			mAP.switchFrame();
			switch (WorkspaceCondition) {
			case "All":
				VerifyAndFetchScreenshots(row);

				try {

					Clib.WaitForElementToLoad(driver, 60, GetWorkSpaceDropDown());
					GetWorkSpaceDropDown().click();
					Clib.StandardWait(2000);
					String wp = driver.findElements(WorkSpace).get(i).getText();
					Clib.StandardWait(2000);
					List<WebElement> SlipperyElement = driver.findElements(WorkSpace);
					size = SlipperyElement.size();

					for (WebElement value : SlipperyElement) {
						Clib.StandardWait(2000);
						if (value.getText().equalsIgnoreCase(wp)) {
							value.click();
							Clib.StandardWait(2000);
							Clib.WaitForElementToLoad(driver, 60, GetDashboard("Top Campaigns"));
							screenshotUtility.TakeScreenshot(GetDashboard("Top Campaigns"), "Top Campaigns_" + wp);
							screenshotUtility.TakeScreenshot(GetDashboard("Top Content"), "Top Content_" + wp);
							screenshotUtility.TakeScreenshot(GetTopIndustries(), "Top Industries_" + wp);
							screenshotUtility.TakeScreenshot(GetTotalOrganizations_Top_Organizations(),
									"Total Organizations_Top Organizations" + wp);
							Clib.WaitForElementToLoad(driver, 60, GetTotalOrgCount());
							Clib.WriteExcelData("Sheet1", TotalOrgCountrow, 0, "Total Organizations");
							Clib.WriteExcelData("Sheet1", TotalOrgCountrow, 1, GetTotalOrgCount().getText());

							Actions actions = new Actions(driver);
							actions.sendKeys(Keys.PAGE_UP).perform();
							actions.sendKeys(Keys.PAGE_UP).perform();
							actions.sendKeys(Keys.PAGE_UP).perform();

							Clib.MouseHover(GetMktoBall());
							Clib.StandardWait(2000);
							GetWebCampaigns().click();
							Clib.StandardWait(2000);

							try {
								GetWebCampaignsNoResults().isDisplayed();
								screenshotUtility.TakeScreenshot(GetCampaignsList(), "Web Campaigns_" + wp);
								break;
							} catch (Exception e) {
								Reporter.log(wp + "No WebCampaigns are present");
								break;
							}

						}

					}
					driver.close();
					driver.switchTo().window(Parent_window);

				}

				catch (Exception e) {
					driver.close();
					driver.switchTo().window(Parent_window);
					break;
				}

			default:
				break;

			case "Specific":
				VerifyAndFetchScreenshots(row);
				try {

					Clib.WaitForElementToLoad(driver, 60, GetWorkSpaceDropDown());
					GetWorkSpaceDropDown().click();
					Clib.StandardWait(2000);
					String wp = prop.getProperty("WorkSpace" + ++i);
					Clib.StandardWait(2000);
					List<WebElement> SlipperyElement = driver.findElements(WorkSpace);
					size = Integer.parseInt(prop.getProperty("NoOfWorkspaces"));

					for (WebElement value : SlipperyElement) {
						Clib.StandardWait(2000);
						if (value.getText().equalsIgnoreCase(wp)) {
							value.click();
							Clib.StandardWait(2000);
							Clib.WaitForElementToLoad(driver, 60, GetDashboard("Top Campaigns"));
							screenshotUtility.TakeScreenshot(GetDashboard("Top Campaigns"), "Top Campaigns_" + wp);
							screenshotUtility.TakeScreenshot(GetDashboard("Top Content"), "Top Content_" + wp);
							screenshotUtility.TakeScreenshot(GetTopIndustries(), "Top Industries_" + wp);
							screenshotUtility.TakeScreenshot(GetTotalOrganizations_Top_Organizations(),
									"Total Organizations_Top Organizations" + wp);
							Clib.WaitForElementToLoad(driver, 60, GetTotalOrgCount());
							Clib.WriteExcelData("Sheet1", TotalOrgCountrow, 0, "Total Organizations");
							Clib.WriteExcelData("Sheet1", TotalOrgCountrow, 1, GetTotalOrgCount().getText());

							Actions actions = new Actions(driver);
							actions.sendKeys(Keys.PAGE_UP).perform();
							actions.sendKeys(Keys.PAGE_UP).perform();
							actions.sendKeys(Keys.PAGE_UP).perform();

							Clib.MouseHover(GetMktoBall());
							Clib.StandardWait(2000);
							GetWebCampaigns().click();
							Clib.StandardWait(2000);

							try {
								GetWebCampaignsNoResults().isDisplayed();
								Reporter.log(wp + "No WebCampaigns are present");
								break;
							} catch (Exception e) {
								screenshotUtility.TakeScreenshot(GetCampaignsList(), "Web Campaigns_" + wp);
								break;
							}
						}

					}
					driver.close();
					driver.switchTo().window(Parent_window);

				}

				catch (Exception e) {
					driver.close();
					driver.switchTo().window(Parent_window);
					break;
				}
			}

		}

	}

}
