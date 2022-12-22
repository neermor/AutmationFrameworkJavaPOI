package com.marketo.qa.utility;

import java.awt.image.BufferedImage;
import java.io.File;
import java.io.IOException;

import javax.imageio.ImageIO;

import org.apache.commons.io.FileUtils;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.Point;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;

import com.assertthat.selenium_shutterbug.core.Capture;
import com.assertthat.selenium_shutterbug.core.Shutterbug;
import com.marketo.qa.base.TestBase;

public class screenshotUtility extends TestBase {

	public static void TakeScreenshot(WebDriver driver, String screenshotName) {
		// take screenshot of the page
		File src = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);
		try {
			FileUtils.copyFile(src,
					new File(System.getProperty("user.dir") + "//Config//ScreenShot//" + screenshotName + ".png"));
		} catch (IOException e) {
			System.out.println("Screenshot Taken" + e.getMessage());
			e.printStackTrace();
		}

	}

	public static void TakeScreenshot(WebElement element, String screenshotName) {
		File src = ((TakesScreenshot) element).getScreenshotAs(OutputType.FILE);
		try {
			FileUtils.copyFile(src,
					new File(System.getProperty("user.dir") + "//Config/ScreenShot//" + screenshotName + ".png"));
		} catch (IOException e) {
			System.out.println("Screenshot Taken" + e.getMessage());
			e.printStackTrace();
		}

	}

	public static void TakeFullPageScreenshot(String screenshotName) {

		Shutterbug.shootPage(driver, Capture.FULL, true).withName(screenshotName)
				.save(System.getProperty("user.dir") + "//Config/ScreenShot//");

	}

	public static void TakeScreenshotForModels(WebElement element, String screenshotName) {
		File src = ((TakesScreenshot) element).getScreenshotAs(OutputType.FILE);
		try {
			FileUtils.copyFile(src, new File(
					System.getProperty("user.dir") + "//Config/ScreenShot/Approved Models/" + screenshotName + ".png"));
		} catch (IOException e) {
			System.out.println("Screenshot Taken" + e.getMessage());
			e.printStackTrace();
		}

	}

	public static void shootWebElement(WebElement element) throws IOException {

		File screen = ((TakesScreenshot) element).getScreenshotAs(OutputType.FILE);

		Point p = element.getLocation();
		System.out.println(p);

		int width = element.getSize().getWidth();
		int height = element.getSize().getHeight();
		System.out.println(width);
		System.out.println(height);

		BufferedImage img = ImageIO.read(screen);

		BufferedImage dest = img.getSubimage(p.getX(), p.getY(), width, height);

		ImageIO.write(dest, "png", screen);

		File f = new File(System.getProperty("user.dir") + "//Config/ScreenShot/");

		FileUtils.copyFile(screen, f);

	}

}
