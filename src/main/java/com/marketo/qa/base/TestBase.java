package com.marketo.qa.base;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Properties;
import java.util.concurrent.TimeUnit;

import org.apache.commons.io.FileUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.interactions.Actions;
import org.testng.annotations.AfterSuite;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.BeforeSuite;

import com.marketo.qa.FileLib.CommonLib;
import com.marketo.qa.Pages.GhostLoginPage;
import com.marketo.qa.Pages.LoginPage;
import com.marketo.qa.Pages.MyMarketoPage;
import com.marketo.qa.utility.ExcelARReport;
import com.marketo.qa.utility.reports;

import io.github.bonigarcia.wdm.WebDriverManager;

public class TestBase {

	private static Logger logger = LogManager.getLogger(TestBase.class);

	public static WebDriver driver;
	public static Properties prop;

	public TestBase() {
		prop = new Properties();
		FileInputStream fis;
		try {
			fis = new FileInputStream(System.getProperty("user.dir") + "//Config//data.properties");
			prop.load(fis);
		} catch (FileNotFoundException e) {

			// TODO Auto-generated catch block
			logger.info(" Data.Properties not fount");
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	@BeforeSuite
	public static void initialization() throws Throwable {

		String browserName = prop.getProperty("browser");

		if (browserName.equals("Chrome")) {

			driver = WebDriverManager.chromedriver().create();

		}

		else if (browserName.equals("Firefox")) {

			driver = WebDriverManager.firefoxdriver().create();
		}

		else if (browserName.equals("IE")) {

			driver = WebDriverManager.iedriver().create();
		}

		else if (browserName.equals("Edge")) {

			driver = WebDriverManager.edgedriver().create();
		} else if (browserName.equals("Safari")) {

			driver = WebDriverManager.safaridriver().create();
		}
		logger.info(browserName + " Browser launch");
		driver.manage().timeouts().implicitlyWait(20, TimeUnit.SECONDS);
		driver.manage().window().maximize();
		driver.manage().deleteAllCookies();
		driver.get(prop.getProperty("ghosturl"));
		new GhostLoginPage().verifyLoginPage();
		// GLogin();
		GhostLogin();
	}

	public static void OpenSupportTool() {
		String url = driver.getCurrentUrl();
		String[] words = url.split("/");
		driver.get(words[0] + "//" + words[2] + "/supportTools");
		logger.info(" supportTools navigating");
	}

	public static void GhostLogin() throws Throwable {
		new GhostLoginPage().GhostLogin(prop.getProperty("prefix"), prop.getProperty("ghostpwd"),
				prop.getProperty("id"));
	}

	public static void GLogin() throws Throwable {
		new GhostLoginPage().GLogin(prop.getProperty("prefix"), prop.getProperty("ghostpwd"), prop.getProperty("id"));
	}

	public static void Login() throws Throwable {
		new LoginPage().Login(prop.getProperty("username"), prop.getProperty("password"));
		new CommonLib().StandardWait(20000);

	}

	public static void Logout() {
		new MyMarketoPage().GetAccountIcon().click();
		new MyMarketoPage().GetLogoutBtn().click();

	}

	@BeforeSuite
	public static void clearScreenshots() {
		String filePath = System.getProperty("user.dir") + "//Config//ScreenShot";
		// Creating the File object
		File file = new File(filePath);
		try {
			FileUtils.deleteDirectory(file);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		logger.info("Screenshot Deleted");
	}

	@AfterTest
	public void BackToMyMarketo() {
		driver.switchTo().defaultContent();
		new MyMarketoPage().OpenMyMarketo();
	}

	@BeforeClass
	public void CloseErrorPopup() {
		for (int i = 1; i < 3; i++) {
			Actions act = new Actions(driver);
			act.sendKeys(Keys.ESCAPE).perform();
		}
	}

	@BeforeMethod
	public void ReturnIframe() {
		driver.switchTo().defaultContent();
	}

	@AfterSuite
	public void teardown() throws Exception {
		String ReportType = prop.getProperty("ReportType");
		switch (ReportType) {
		case "IR":
			reports.docs();
			break;

		case "AR":
			ExcelARReport.ARReport();
			break;

		case "Both":
			reports.docs();
			ExcelARReport.ARReport();
			break;

		default:
			break;
		}

		driver.quit();
		System.out.println("Execution Over");

	}
}
