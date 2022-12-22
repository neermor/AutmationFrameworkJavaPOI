package com.marketo.qa.utility;

import java.io.File;
import java.io.IOException;


import org.apache.commons.io.FileUtils;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;



public class screenshotUtility {
	
	 public static void TakeScreenshot(WebDriver driver, String screenshotName)
	    {
	        //take screenshot of the page
	        File src= ((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
	        try {
	            FileUtils.copyFile(src, new File("C:\\Users\\neeraj.mourya\\eclipse-workspace\\Acedmy\\test-output\\screenshots\\"+screenshotName+".png"));
	        } catch (IOException e) {
	            System.out.println("Screenshot Taken"+e.getMessage());
	            e.printStackTrace();
	        }
		
	    }
	 
	 public static void TakeScreenshot(WebElement element, String ScreenshotName)
		{
			File src = ((TakesScreenshot) element).getScreenshotAs(OutputType.FILE);
			try {
				FileUtils.copyFile(src,
						new File(System.getProperty("user.dir") + "\\test-output\\screenshots\\" + ScreenshotName + ".png"));
			} catch (IOException e) {
				System.out.println("Screenshot Taken" + e.getMessage());
				e.printStackTrace();
			}
			
		}

}
