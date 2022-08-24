package com.marketo.qa.Pages;

import java.util.List;

import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;

import com.marketo.qa.FileLib.CommonLib;
import com.marketo.qa.base.TestBase;
import com.marketo.qa.utility.screenshotUtility;

import net.lightbody.bmp.proxy.jetty.html.Break;

public class AnalyticsPage extends TestBase{
	MyMarketoPage homepage= new MyMarketoPage();

	By Models=By.xpath("//div[contains(@data-id,'treeNode_revenuecyclemodel')]");
	By Rcm = By.cssSelector("#treeBodyAnchor > div > div > div:nth-child(4)");
	
	public List<WebElement> GetModels() {
		return driver.findElements(Models);
	}
	public WebElement GetRcm() {
		return driver.findElement(Rcm);
	}
	

		public void ModelCount(int row) throws Throwable {
			homepage.ExtendTreeNode("Revenue Cycle Modeler");	
			try {
				boolean  flag=driver.findElement(By.xpath("//div[contains(@data-id,'treeNode_Label')]/span[text()='Group Models']/../preceding-sibling::button[@data-id='treeNodeChevronIconButton']")).isDisplayed();
				if(flag) {
					homepage.ExtendTreeNode("Group Models");
					new CommonLib().WriteExcelData("Sheet1", row, 0, "Models");
					System.out.println(GetModels().size());
					new CommonLib().WriteExcelData("Sheet1", row, 1, GetModels().size());
					screenshotUtility.TakeScreenshot(GetRcm(), "Models");		 		
				}
			}
			catch (Exception e) {
				new CommonLib().WriteExcelData("Sheet1", row, 0, "Models");
				System.out.println(GetModels().size());
				new CommonLib().WriteExcelData("Sheet1", row, 1, 0);
				}
			
			
			}

}