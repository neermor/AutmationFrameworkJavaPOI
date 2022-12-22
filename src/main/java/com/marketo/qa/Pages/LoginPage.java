package com.marketo.qa.Pages;

import java.util.concurrent.TimeUnit;

import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;

import com.marketo.qa.base.TestBase;

public class LoginPage extends TestBase{
	

	private WebElement UserName=driver.findElement(By.name("username"));
	private WebElement Password= driver.findElement(By.name("passwd"));
	private WebElement LoginButton= driver.findElement(By.cssSelector("#loginButton"));

	
	public WebElement getUserName() {
		return UserName;
	}

	public WebElement getPassword() {
		return Password;
	}

	public WebElement getLoginButton() {
		return LoginButton;
	}
	
	public MyMarketoPage Login(String un, String pwd) {
		driver.manage().timeouts().implicitlyWait(40, TimeUnit.SECONDS);
		UserName.sendKeys(un);
		Password.sendKeys(pwd);
		LoginButton.click();
		return new MyMarketoPage();
	}
	

}
