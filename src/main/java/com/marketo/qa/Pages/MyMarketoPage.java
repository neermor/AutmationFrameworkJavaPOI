package com.marketo.qa.Pages;

import java.util.List;
import java.util.concurrent.TimeUnit;

import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;
import org.testng.Assert;

import com.marketo.qa.base.TestBase;


public class MyMarketoPage extends TestBase {
	
	boolean flag = false;
	By MyMarketoPageHomeTiles=By.xpath("//div[@role='tablist']/div");
	By AccountIcon= By.cssSelector("div [data-id='userAccountNavButton']");
	By LogoutBtn = By.cssSelector("[data-id='userAccountLogoutButton']");
	By TreeNode=By.cssSelector("[data-id='globalTreeDrawerExpanderContent'] [data-id='treeNode_Label']");

	public WebElement GetAccountIcon() {
		return driver.findElement(AccountIcon);
	}
	
	public WebElement GetLogoutBtn() {
		return driver.findElement(LogoutBtn);
	}
	
	public void OpenMarketingActivitiesTab() throws Throwable {
		driver.manage().timeouts().implicitlyWait(40, TimeUnit.SECONDS);

	    List<WebElement> HomeTiles = driver.findElements(MyMarketoPageHomeTiles);
        System.out.println(HomeTiles);
		 for(WebElement option: HomeTiles){
			 
				if (option.getText().equalsIgnoreCase("Marketing Activities")) {
					option.click();
				}
				System.out.println(option.getText());
				Thread.sleep(4000);
			 }
		 }
	
	public void OpendesignStudioTab() {
	   List<WebElement> HomeTiles = driver.findElements(MyMarketoPageHomeTiles);

		driver.manage().timeouts().implicitlyWait(40, TimeUnit.SECONDS);		
		 for(WebElement option: HomeTiles){
			 
				if (option.getText().equalsIgnoreCase("design Studio")) {
					option.click();
				}
			 }
		 }
	
	public void OpenAdminTab() {
		 List<WebElement> HomeTiles = driver.findElements(MyMarketoPageHomeTiles);

			driver.manage().timeouts().implicitlyWait(40, TimeUnit.SECONDS);		
			 for(WebElement option: HomeTiles){
				 
					if (option.getText().equalsIgnoreCase("admin")) {
						option.click();
					}
				 }
			 }
	
	public void OpenanalyticsTab() {
		 List<WebElement> HomeTiles = driver.findElements(MyMarketoPageHomeTiles);

			driver.manage().timeouts().implicitlyWait(40, TimeUnit.SECONDS);		
			 for(WebElement option: HomeTiles){
				 
					if (option.getText().equalsIgnoreCase("analytics")) {
						option.click();
					}
				 }
			 }
	
	public void OpenDatabaseTab() {
		 List<WebElement> HomeTiles = driver.findElements(MyMarketoPageHomeTiles);

			driver.manage().timeouts().implicitlyWait(40, TimeUnit.SECONDS);		
			 for(WebElement option: HomeTiles){
				 
					if (option.getText().equalsIgnoreCase("database")) {
						option.click();
					}
				 }
			 }
	
	
	public boolean VerifyHomeTileElements() throws Throwable {
		Thread.sleep(10000);

	    List<WebElement> HomeTiles = driver.findElements(MyMarketoPageHomeTiles);

	    
	    if (HomeTiles.size() == 0) {
	    	Assert.assertTrue(false, "No HomeTiles were found on the page!: ");
	    	}
	    else {
	        for(WebElement option: HomeTiles){
	        	Assert.assertTrue(option.isEnabled(),"HomeTiles were found on the page!:");
	       flag = true;
	        	
	         }	        
	    }
		return flag;
	}

	public void SelectTreeNode(String TreeNodeName) throws Throwable {
		Thread.sleep(10000);
	    List<WebElement> HomeTiles = driver.findElements(TreeNode);
        System.out.println(HomeTiles);
		 for(WebElement option: HomeTiles){
			 
				if (option.getText().equalsIgnoreCase(TreeNodeName)) {
					option.click();
				}
				System.out.println(option.getText());
			 }
		 }
	
	public void SelectWorkSpace(String WorkspaceName) throws Throwable {
		SelectTreeNode(WorkspaceName);

	}
	

	public void  ExtendTreeNode(String Name) throws Throwable {
		driver.findElement(By.xpath("//div[contains(@data-id,'treeNode_Label')]/span[text()="+"'"+Name+"'"+"]/../preceding-sibling::button[@data-id='treeNodeChevronIconButton']")).click();
	}
	
	

}
