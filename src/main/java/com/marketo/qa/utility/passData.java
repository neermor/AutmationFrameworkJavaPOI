package com.marketo.qa.utility;

import java.io.IOException;
import java.util.Map;






public class passData  {
	
	
	
	
	public static String FetchScreenshot(String ScreenshotName) {
		 String path = "C:\\Users\\pradyumna.sahoo\\eclipse-workspace\\Acedmy\\test-output\\screenshots\\"+ScreenshotName+".png";
		return path;
	}
	
	static String Tag = "C:\\Users\\pradyumna.sahoo\\Dev\\MarketoInstance\\test-output\\screenshots\\Tags.png";
	static String intesting_moment_img = "C:\\Users\\neeraj.mourya\\eclipse-workspace\\MarketoInstance\\MarketoInstance\\test-output\\screenshots\\Interesting Moment.png";
	
	static String Document_name = "8967";
	static String Org_info = " has created a lot of campaigns and content in Marketo."
			+ " There are two types of smart campaigns: Batch and Trigger. A batch campaign launches "
			+ "at a specific time and affects a specific set of leads all at once. A triggered smart"
			+ " campaign affects one lead at a time, based on a triggered event."
			+ " To learn more about Smart campaigns in Marketo. visit: https://docs.marketo.com/display/public/DOCS/Smart+Campaigns";
	
	static String models  = "has built models in Marketo. The “Shared Revenue Models” folder is shared across the following workspaces:"
			+ " BizDev, Buyers Journey, Customer Journey, and DemandGen (content). Revenue cycle models take marketing to the next level. "
			+ "They model all the stages of your entire revenue funnel—from when you first interact with a lead all the way until the lead"
			+ " is a won customer. Infusion Software, Inc."
			+ " has created the following Revenue Cycle Models in their Marketo instance:";
	
	static String lead_scoring ="Marketo’s lead scoring capabilities are far more robust than any other vendor offerings. \r\n"
			+ "Lead scoring allows you to identify which prospects are most interested and engaged with your brand. "
			+ "Marketo also allows the usage of My Tokens in lead scoring campaigns. This allows the marketer to have "
			+ "the ability to control at a high level all of the lead scoring attributes assigned to their campaigns."
			+ " Additionally, Marketo allows the marketer to add detailed constraints to their lead scoring campaigns, "
			+ "which add another layer of complexity. For example – leads active during a specific date/time AND who "
			+ "visit the web page numerous times within a certain time window.";
	
	static String intresting_moment = "The following Interesting Moments have been defined to support";
	
	static String intresting_moment_below ="If you have Marketo Sales Insight, you can use the interesting moment flow step to give your"
			+ " sales team visibility into the cool things your leads are doing in a Smart Campaign. Interesting Moments allow the marketer "
			+ "to define what information is relevant to their sales team. "
			+ "When a lead takes a specific action, that action is logged and recorded for the team to see."; 
	
	static String data_management ="has Data Management actions setup within Marketo. Marketo’s data management tools allow a marketer to configure "
			+ "actions to automatically manage leads. For example, Data Management actions can be set up to de-duplicate data,"
			+ " clean up bad data, and modify data based on predetermined actions and values.";
	
	static String events = "One of the greatest features of Marketo is the ability to clone an entire program—which copies"
			+ " all underlying assets and campaigns that are part of that program. Events allow you to automate online and "
			+ "offline events! Capture the status of your leads as they progress through different stages and get accurate"
			+ " measures of the ROI for your marketing initiatives.";
	
	static String Nurture= " Inc. has built 7 Nurture campaigns using the Marketo Nurture Stream engine. There are two types"
			+ " of Content you can add to engagement program streams — emails and programs. Emails will be sent to leads at cast "
			+ "time. Marketo's smart streams also offer:";
	
	public static String Exceldata(String value) throws IOException 
	{
		Map <String, String> testdata =excelReader.getMapData();
		
		return testdata.get(value);
		
	}
	
}
	
	
	
	
	

