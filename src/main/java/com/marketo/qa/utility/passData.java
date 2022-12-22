package com.marketo.qa.utility;

import java.io.IOException;
import java.util.Map;


import com.marketo.qa.Pages.SupportToolsPage;



public class passData  {
	
	
	
	static SupportToolsPage STP= new SupportToolsPage();
	static String test=String.valueOf(STP.GetCount());
		  
	static String login = "C:\\Users\\neeraj.mourya\\eclipse-workspace\\Acedmy\\test-output\\screenshots\\login.png";
	static String after = "C:\\Users\\neeraj.mourya\\eclipse-workspace\\Acedmy\\test-output\\screenshots\\after.png";
	
	static String Document_name = "8967";
	static String Org_info = "AvePoint_Inc has created a lot of campaigns and content in Marketo."
			+ " There are two types of smart campaigns: Batch and Trigger. A batch campaign launches "
			+ "at a specific time and affects a specific set of leads all at once. A triggered smart"
			+ " campaign affects one lead at a time, based on a triggered event."
			+ " To learn more about Smart campaigns in Marketo. visit:https://docs.marketo.com/display/public/DOCS/Smart+Campaigns .";
			
	public static String Exceldata(String value) throws IOException 
	{
		Map <String, String> testdata =excelReader.getMapData();
		
		return testdata.get(value);
		
	}
	 
	

	
	
}
	
	
	
	
	

