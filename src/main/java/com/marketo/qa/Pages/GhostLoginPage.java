package com.marketo.qa.Pages;

import java.util.concurrent.TimeUnit;
import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;

import com.marketo.qa.base.TestBase;


public class GhostLoginPage extends TestBase{
	

	private By Password= By.id("loginPassword");
	private By GhostId= By.id("secondaryUsername");
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
	
	public void GhostLogin(String prefix, String pwd, String ghostId) throws Throwable {
		driver.manage().timeouts().implicitlyWait(40, TimeUnit.SECONDS);
		GetPrfix().sendKeys(prefix);
		getPassword().sendKeys(pwd);
		getGhostID().clear();
		getGhostID().sendKeys(ghostId);
		getLoginButton().click();
		
		if(GetOktaUsername().isDisplayed()) {
			Thread.sleep(2000);	
			GetOktaUsername().sendKeys(prop.getProperty("OcktaID"));
			GetOktaNextBtn().click();
			Thread.sleep(2000);	
			GetOktaPassword().sendKeys(prop.getProperty("OcktaPass"));
			GetOktaVerifyBtn().click();	
			Thread.sleep(2000);		
			GetOktaPushBtn().click();
		}
	}
	

}
