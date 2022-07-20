package com.marketo.qa.Pages;

import java.util.concurrent.TimeUnit;

import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;

import com.marketo.qa.base.TestBase;

public class GhostLoginPage extends TestBase{
	

	private WebElement UserName=driver.findElement(By.name("username"));
	private WebElement Password= driver.findElement(By.name("passwd"));
	private WebElement LoginButton= driver.findElement(By.cssSelector("#loginButton"));
	private By prifix = By.cssSelector("loginUsername");
	
	public WebElement getUserName() {
		return UserName;
	}

	public WebElement GetPrfix() {
		return driver.findElement(prifix);
	}
	public WebElement getPassword() {
		return Password;
	}

	public WebElement getLoginButton() {
		return LoginButton;
	}
	
	public MyMarketoPage GhostLogin(String un, String pwd) {
		driver.manage().timeouts().implicitlyWait(40, TimeUnit.SECONDS);
		GetPrfix().sendKeys(un);
		Password.sendKeys(pwd);
		LoginButton.click();
		return new MyMarketoPage();
	}
	

}
