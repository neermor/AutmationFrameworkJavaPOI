package com.marketo.qa.Pages;

import java.util.concurrent.TimeUnit;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;
import org.testng.Assert;

import com.marketo.qa.FileLib.CommonLib;
import com.marketo.qa.base.TestBase;

public class GhostLoginPage extends TestBase {
	private static Logger logger = LogManager.getLogger(TestBase.class);
	CommonLib Clib = new CommonLib();
	MyMarketoPage homePage = new MyMarketoPage();

	private By Password = By.id("loginPassword");
	private By GhostId = By.id("secondaryUsername");
	private By Prifix = By.id("loginUsername");
	private By Signup = By.id("loginButton");
	private By OkatUsername = By.cssSelector("#okta-signin-username");
	private By OkatNextBtn = By.cssSelector("#okta-signin-submit");
	private By OkatPassword = By.cssSelector("[name='password']");
	private By OkatVerifyBtn = By.cssSelector("[value='Verify']");
	private By OkatSendPush = By.cssSelector("[value='Send Push']");
	private By OkatSPushsent = By.cssSelector("[value='Push sent!']");
	private By OktaWarning = By.cssSelector("icon warning");

	public WebElement GetPrfix() {
		return driver.findElement(Prifix);
	}

	public WebElement GetOktaWarning() {
		return driver.findElement(OktaWarning);
	}

	public WebElement getPassword() {
		return driver.findElement(Password);
	}

	public WebElement getLoginButton() {
		return driver.findElement(Signup);
	}

	public WebElement getGhostID() {
		return driver.findElement(GhostId);
	}

	public WebElement GetOktaUsername() {
		return driver.findElement(OkatUsername);
	}

	public WebElement GetOktaNextBtn() {
		return driver.findElement(OkatNextBtn);
	}

	public WebElement GetOktaPassword() {
		return driver.findElement(OkatPassword);
	}

	public WebElement GetOktaVerifyBtn() {
		return driver.findElement(OkatVerifyBtn);
	}

	public WebElement GetOktaPushBtn() {
		return driver.findElement(OkatSendPush);
	}

	public WebElement GetOktaPushSent() {
		return driver.findElement(OkatSPushsent);
	}

	public boolean verifyLoginPage() {
		boolean flag = false;
		try {
			flag = GetPrfix().isDisplayed();

		} catch (Exception e) {
			logger.info("Login Page is Not dispalyed !! Kindly Check once your Internet connection");
		}
		return flag;
	}

	public boolean OktaVerify() {
		boolean flag = false;
		String titel = driver.getTitle();
		System.out.println(titel);
		if (titel.equalsIgnoreCase("Adobe Marketo Engage") || titel.equalsIgnoreCase("My Marketo")) {
			flag = true;
		}

		return flag;
	}

	public void OctaLogs() {
		if (OktaVerify()) {
			logger.info("Okta verified Sucessfully");
		} else {
			logger.info("Failed Okta verified");
		}
	}

	int count = 0;

	public void GhostLogin(String prefix, String pwd, String ghostId) throws Throwable {
		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
		GetPrfix().sendKeys(prefix);
		getPassword().sendKeys(pwd);
		getGhostID().clear();
		getGhostID().sendKeys(ghostId);
		getLoginButton().click();

		boolean push = false;

		try {
			push = GetOktaPushBtn().isDisplayed();
			if (push) {
				GetOktaPushBtn().click();
				new CommonLib().StandardWait(40000);
				OctaLogs();
				Assert.assertTrue(OktaVerify());
			}

		} catch (Exception e) {
			logger.info("Direct Okta Push Button Is Not Presnet");
		}
		while (!push) {

			try {
				GetOktaUsername().isDisplayed();
				Thread.sleep(2000);
				GetOktaUsername().sendKeys(prop.getProperty("OcktaUserID"));
				logger.info("Entered Okta Username");
				GetOktaNextBtn().click();
				Thread.sleep(2000);
				GetOktaPassword().sendKeys(prop.getProperty("OcktaPassword"));
				logger.info("Entered Okta Password");
				GetOktaVerifyBtn().click();
				Thread.sleep(2000);
				GetOktaPushBtn().click();
				logger.info("Clicked Push Sent Button");
				new CommonLib().StandardWait(40000);
				OctaLogs();
				Assert.assertTrue(OktaVerify());
				break;
			}

			catch (Exception e) {
			}
		}

		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
	}

}
