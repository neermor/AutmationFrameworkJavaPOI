package com.marketo.qa.Pages;

import java.util.List;

import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;

import com.marketo.qa.FileLib.CommonLib;
import com.marketo.qa.base.TestBase;

public class DesignStudioPage extends TestBase{
	
	By TreeNode=By.cssSelector("[data-id='globalTreeDrawerExpanderContent'] [data-id='treeNode_Label']");
	By AllCount = By.cssSelector("[data-id='dataGridFooter_pageInfo']");
	
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

	public String GetCount() throws Throwable {
		Thread.sleep(4000);
	    String countString = driver.findElement(AllCount).getText();
	    String[] words=countString.split("\\s");  
	    System.out.println(words[4]); 
	    return words[4];
	}
	
	public void FetchTreeNodeCount(String value,int row) throws Throwable {
		SelectTreeNode(value);
		Thread.sleep(2000);
		new CommonLib().WriteExcelData("Sheet1", row, 0, value);
		new CommonLib().WriteExcelData("Sheet1", row, 1, GetCount());	

	}
	
	

}
