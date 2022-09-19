package com.marketo.qa.Pages;

import com.marketo.qa.FileLib.CommonLib;
import com.marketo.qa.base.TestBase;

public class PredictiveContentPage extends TestBase{
	MyMarketoPage homepage= new MyMarketoPage(); 
		
		public void CheckPresentOrNot(int row) throws Exception {
			try {
			homepage.GetHometileUnderFrame("Predictive Content").isDisplayed();

			new CommonLib().WriteExcelData("Sheet1", row, 0, "Predictive Content");
			new CommonLib().WriteExcelData("Sheet1", row, 1, "True");			
			}
			
			catch (Exception e) {
				new CommonLib().WriteExcelData("Sheet1", row, 0, "Predictive Content");
				new CommonLib().WriteExcelData("Sheet1", row, 1, "False");
			}
			
			
		}
	}
