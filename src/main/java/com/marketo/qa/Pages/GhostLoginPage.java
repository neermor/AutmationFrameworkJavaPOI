package com.marketo.qa.Pages;

import java.util.concurrent.TimeUnit;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;

import com.marketo.qa.base.TestBase;

public class GhostLoginPage extends TestBase {
	private static Logger logger = LogManager.getLogger(TestBase.class);
	
	
	private By Password = By.id("loginPassword");
	private By GhostId = By.id("secondaryUsername");
	private By Prifix = By.id("loginUsername");
	private By Signup = By.id("loginButton");
	private By OkatUsername = By.cssSelector("#okta-signin-username");
	private By OkatNextBtn = By.cssSelector("#okta-signin-submit");
	private By OkatPassword = By.cssSelector("[name='password']");
	private By OkatVerifyBtn = By.cssSelector("[value='Verify']");
	private By OkatSendPush = By.cssSelector("[value='Send Push']");

	public WebElement GetPrfix() {
		return driver.findElement(Prifix);
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

	int count = 0;

	public void GhostLogin(String prefix, String pwd, String ghostId) throws Throwable {
		driver.manage().timeouts().implicitlyWait(40, TimeUnit.SECONDS);
		GetPrfix().sendKeys(prefix);
		getPassword().sendKeys(pwd);
		getGhostID().clear();
		getGhostID().sendKeys(ghostId);
		getLoginButton().click();

		boolean push = false;

		driver.manage().timeouts().implicitlyWait(1, TimeUnit.SECONDS);
		int flag = 0;
		while ((driver.findElements(OkatSendPush).size() > 0)
				|| (driver.findElements(OkatUsername).size() > 0) && flag < 3) {
			Thread.sleep(500);
			flag++;
		}

		try {
			push = GetOktaPushBtn().isDisplayed();
			if (push) {
				GetOktaPushBtn().click();
				logger.info("Okta verified");
			}

		} catch (Exception e) {
			logger.info("Direct Okta Push Button Is Not Presnet");
		}
		while (!push) {

			try {
				GetOktaUsername().isDisplayed();
				Thread.sleep(2000);
				GetOktaUsername().sendKeys(prop.getProperty("OcktaUserID"));
				GetOktaNextBtn().click();
				Thread.sleep(2000);
				GetOktaPassword().sendKeys(prop.getProperty("OcktaPassword"));
				GetOktaVerifyBtn().click();
				Thread.sleep(2000);
				GetOktaPushBtn().click();
				logger.info("Okta verified");

				break;
			}

			catch (Exception e) {
				logger.info("Failed Okta verified");
			}
		}
		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
	}

}
