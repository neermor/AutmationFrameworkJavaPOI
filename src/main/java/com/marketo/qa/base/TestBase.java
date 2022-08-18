package com.marketo.qa.base;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Properties;
import java.util.concurrent.TimeUnit;

import org.openqa.selenium.WebDriver;
import org.testng.Reporter;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;

import com.marketo.qa.FileLib.CommonLib;
import com.marketo.qa.Pages.GhostLoginPage;
import com.marketo.qa.Pages.LoginPage;
import com.marketo.qa.Pages.MyMarketoPage;

import io.github.bonigarcia.wdm.WebDriverManager;


public class TestBase {
	public static WebDriver driver;
	public static Properties prop;

	public TestBase() {
		prop = new Properties();

		FileInputStream fis;
		try {
			fis = new FileInputStream(System.getProperty("user.home") + "/Desktop/Config/data.properties");
			prop.load(fis);

		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}


	@BeforeTest
	public static void initialization() {

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
		}
		else if (browserName.equals("Safari")) {
		
			driver = WebDriverManager.safaridriver().create();
		}
		
		driver.manage().timeouts().implicitlyWait(20, TimeUnit.SECONDS);
		driver.manage().window().maximize();
		driver.manage().deleteAllCookies();
		driver.get(prop.getProperty("ghosturl"));
	}

	public static void OpenSupportTool() {
		driver.get(prop.getProperty("url") + "/supportTools");
	}


	
	public static void GhostLogin() throws Throwable {
		new GhostLoginPage().GhostLogin(prop.getProperty("prifix"), prop.getProperty("Ghostpwd"),prop.getProperty("id"));
		new CommonLib().StandardWait(20000);

	}
	public static void MarketoLogin() throws Throwable {
		new LoginPage().Login(prop.getProperty("username"), prop.getProperty("password"));
		new CommonLib().StandardWait(20000);
		Reporter.log("Login Succsessfully");
	}
	
	public static void Logout() {
		new MyMarketoPage().GetAccountIcon().click();
		new MyMarketoPage().GetLogoutBtn().click();

	}
	

	  public static void CloseBrowser() { 
		  driver.quit();
		  }
	 

}
