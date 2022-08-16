package com.marketo.qa.utility;


import java.io.IOException;
import java.util.Map;

import com.marketo.qa.Pages.AdminPage;






public class passData  {
	static AdminPage Ap = new AdminPage();
	static String AccountName= Ap.AccountName();
	
	public static String FetchScreenshot(String screenshotName) {
		
		 String path = System.getProperty("user.home") + "\\Desktop\\Config\\"+screenshotName+".png";
		return path;
	}
	
	
	
	
	static String Org_info = " has created a lot of campaigns and content in Marketo."
			+ " There are two types of smart campaigns: Batch and Trigger. A batch campaign launches "
			+ "at a specific time and affects a specific set of leads all at once. A triggered smart"
			+ " campaign affects one lead at a time, based on a triggered event."
			+ " To learn more about Smart campaigns in Marketo. visit: https://docs.marketo.com/display/public/DOCS/Smart+Campaigns.";
	
	static String No_models  =AccountName  +"\nhas not built models in Marketo. Revenue cycle models take marketing to the next level."
			+ " They model all the stages of your entire revenue funnel—from when you first interact with a lead all "
			+ "the way until the lead is a won customer.";

	static String models  = AccountName +"\nbuilt models in Marketo. Revenue cycle models take marketing to the next level. They model all the stages"
			+ " of your entire revenue funnel—from when you first interact with a lead all the way until the lead is"
			+ " a won customer.";
	
	
	static String models2= AccountName +"\nhas created the following Revenue Cycle Models in their Marketo instance:";
	
	static String lead_scoring ="Marketo’s lead scoring capabilities are far more robust than any other vendor offerings. \r\n"
			+ "Lead scoring allows you to identify which prospects are most interested and engaged with your brand. "
			+ "Marketo also allows the usage of My Tokens in lead scoring campaigns. This allows the marketer to have "
			+ "the ability to control at a high level all of the lead scoring attributes assigned to their campaigns."
			+ " Additionally, Marketo allows the marketer to add detailed constraints to their lead scoring campaigns, "
			+ "which add another layer of complexity. For example – leads active during a specific date/time AND who "
			+ "visit the web page numerous times within a certain time window.";
	
	//Interesting Moment Data 
	
	static String intresting_moment = "The following Interesting Moments have been defined to support\n"+ AccountName  +"’s marketing campaigns."
			+" When a lead exhibits any of the below behavior, it will be documented and tracked.";
	
	static String intresting_moment_below ="If you have Marketo Sales Insight, you can use the interesting moment flow step to give your"
			+ " sales team visibility into the cool things your leads are doing in a Smart Campaign. Interesting Moments allow the marketer "
			+ "to define what information is relevant to their sales team. "
			+ "When a lead takes a specific action, that action is logged and recorded for the team to see.";
	
	static String intresting_moment_else= passData.AccountName+"\nhas not defined any Interesting Moments to support their marketing campaigns in Marketo. Interesting Moments are used to document and track any lead behaviors exhibited that are of importance to the Sales Team and the Marketing Team.\r\n"
			+ "Interesting Moments allow the marketer to define what information is relevant to their sales team. When a lead takes a specific action, that action is logged and recorded for the team to see. There is nothing comparable to interesting moments with other vendors.";
	
	
	//Event Data 
	
	static String events = "\nOne of the greatest features of Marketo is the ability to clone an entire program—which copies"
			+ " all underlying assets and campaigns that are part of that program. Events allow you to automate online and "
			+ "offline events! Capture the status of your leads as they progress through different stages and get accurate"
			+ " measures of the ROI for your marketing initiatives.";
	
	static String No_events = AccountName +"\n does not have any Event campaigns set up in their instance. Here is a good doc to share"
			+ " on how to create a new event campaign:"
			+ " https://docs.marketo.com/display/public/DOCS/Create+a+New+Event+Program";
	
	
	
	//Nurture Data 
	
	static String Nurture= "\nNurture campaigns using the Marketo Nurture Stream engine. There are two "
			+ "types of Content you can add to engagement program streams — emails and programs. Emails will be sent"
			+ "to leads at cast time. Marketo's smart streams also offer:";

	static String No_Nurture= AccountName +"\ndoes not have any Nurture campaigns set up in their instance. They might need a refresher on how"
			+ " useful Nurture Campaigns can be: "
			+ "https://experienceleague.adobe.com/docs/marketo/using/product-docs/email-marketing/drip-nurturing/creating-an-engagement-program/create-an-engagement-program.html?lang=en";
	
	
	//Segments 
	static String no_segment= AccountName  +"\ndoes not have any segments defined. Segmentation categorizes your audience into different"
			+ " subgroups based on a Smart List rule. These groups are called segments. Segments allow the marketer"
			+ " to target leads based on the segment that they fall into.\n";
		
	static String segment= AccountName +"\nhas the following segments defined. Segments allow the marketer to target leads based on the"
			+ " segment that they fall into.";
	
	//Snippet
	static String No_Snippets =  AccountName +"\n is not taking advantage of snippets. Snippets are reusable blocks of rich text and graphics that"
			+ " the client can use in their emails and landing pages, and it is a great timesaver!"
			+ " https://experienceleague.adobe.com/docs/marketo/using/product-docs/personalization/segmentation-and-snippets/snippets/create-a-snippet.html?lang=en";
	
	
	//No Lead_scoring Data
	static String Lead_scoring= "\n has advanced lead scoring built out taking into account behavior, demographics, successes and decay.";
	static String No_lead_scoring= AccountName +"\n has not built out any lead scoring campaigns in Marketo";
	
	static String tokens= AccountName +"\nis using MyTokens in their lead scoring campaigns which allows for a Marketer to quickly, "
			+ "and easily, control from a high level their lead change scores";
	
	static String no_tokens= AccountName +"\nis Not using MyTokens in their lead scoring campaigns which allows for a Marketer to quickly, "
			+ "and easily, control from a high level their lead change scores";
	
	static String no_batch=AccountName +"\nhas Not built out campaigns reducing lead scores when leads exhibit undesirable behavior";
	
	static String batch= "\nbuilt out campaigns reducing lead scores when leads exhibit undesirable behavior";
	
	//Data Management 
	static String Data= AccountName +"\nhas Data Management actions setup within Marketo. Marketo’s data management tools allow a"
			+ " marketer to configure actions to automatically manage leads. For example, Data Management actions can"
			+ " be set up to de-duplicate data, clean up bad data, and modify data based on predetermined actions and "
			+ "values.";
	
	static String No_Data= AccountName +"\ndoes not appear to have data management actions set up. Good examples of data management "
			+ "would be any steps taken to clean up lead data, for example, adding leads to a blacklist triggered by"
			+ " a certain action. Here is a walkthrough of how to do that: https://experienceleague.adobe.com/docs/marketo/using/product-docs/core-marketo-concepts/smart-lists-and-static-lists/managing-people-in-smart-lists/add-person-to-blocklist.html?lang=en";
	static String No_data_link= "\nHere is a high overview on how to create Change Data Value flow actions: ";
	
	//Library Data
	
	static String Program_library= "It appears that\n" +AccountName + "\nhas imported templates from the Marketo Program Library.\r\n"
			+ "Marketo is committed to our customers' success and has seeded a ton of pre-built programs for almost any use case"
			+ " into the Marketo Program Library that our customers are free to import when they are needed as their marketing "
			+ "strategies evolve and call for different types of programs and campaigns.";
	
	static String no_library= "It appears that\n"+AccountName +"\nhas Not imported templates from the Marketo Program Library.\r\n"
			+ "Marketo is committed to our customers' success and has seeded a ton of pre-built programs for almost any use "
			+ "case into the Marketo Program Library that our customers are free to import when they are needed as"
			+ " their marketing strategies evolve and call for different types of programs and campaigns.";
	
	// Integrations 
	
	static String No_Integrations= AccountName +"\nhas Not installed any integrations."
			+ "Marketo LaunchPoint is the most complete ecosystem of Marketing solution integrations in the industry."
			+ " LaunchPoint offers hundreds of applications that complement and integrate into Marketo’s customer"
			+ " engagement platform. LaunchPoint gives Marketers access to the best applications, solutions and"
			+ " services that drive engagement and build revenue.";
	
	//web personalize 
	
	static String web_personalize= "With Marketo’s Real Time Personalization, you can engage the 98% of visitors that"
			+ " you don’t know. 98% of website visitors are anonymous, only 2% are known. Using RTP, Marketers have an "
			+ "opportunity to engage these anonymous visitors with relevant content and personalized messages via web "
			+ "or mobile using firmographic and behavioral data. This real time personalization capability is "
			+ "completely unified within the platform and shares a common user experience, making it easy for marketers"
			+ " to create personalized web experiences without IT support. This results in increased conversion rates"
			+ " up to 30% and increased content engagement up to 270%. Competitors cannot generate personalized web"
			+ " experiences for anonymous visitors, and personalization for known visitors is overly complicated";
	
	//Target Account Management
	static String Account_Management ="Marketo Account Based Marketing brings sales and marketing teams together to"
			+ " target and engage key accounts in a highly coordinated fashion, bridging the gap between"
			+ " account-centric strategy, execution and success – all within a single platform.";
	
	static String Account_Management2 ="Marketo unifies ABM and lead management in one solution, making it easy for"
			+ " marketers to execute personalized campaigns for both accounts and leads in one motion. You also "
			+ "benefit from reaching key decision makers and deal influencers.";
	
	public static String Exceldata(String value) throws IOException 
	{
		Map <String, String> testdata =excelReader.getMapData();
		
		return testdata.get(value);
		
	}	 
	
	
	public static String GenricMethod(String snippets) throws IOException {
		
		int snippet = Integer.parseInt(passData.Exceldata("Snippets"));
		
		
		if (snippet>=5) 
		{		
			
			return snippets+"\nSnippets" ;
		}
		
		else  {
		}
			return  No_Snippets;
			
					
	} 

	//No Lead_scoring Data
public static String LeadScoring(String leads) throws IOException {
		
		int lead = Integer.parseInt(passData.Exceldata("Leads"));
		
		
		if (lead>=5) 
		{			
			return AccountName +"\nhas\n"+leads+ Lead_scoring;
		}
		
		else  {
		}
			return No_lead_scoring;			
	} 

//my tokens method for Lead scoring

public static String MyTokens() throws IOException {	
	int token = Integer.parseInt(passData.Exceldata("Tokens"));
	
	if (token>0) 
	{			
		return tokens ;
	}
	
	else  {
	}
	
		return no_tokens;
				
} 


public static String BatchCampaigns() throws IOException {	
	
	int Batch = Integer.parseInt(passData.Exceldata("All Batch Campaigns"));
	
	if (Batch>0) 
	{			
		return AccountName +"\nhas built\n"+Exceldata("All Batch Campaigns")+batch ;
	}
	
	else  {
	}
	
		return no_batch;
				
} 

//Method for Data Management 

public static String DataManagement() throws IOException {	
	
	int DataManagment = Integer.parseInt(passData.Exceldata("Change Data Value"));
	
	if (DataManagment>10) 
	{			
		return Data ;
	}
	
	else  {
	}
	
		return No_Data +No_data_link ;
				
} 

//Method for Events 
public static String Events() throws IOException {	
	
	int Event_Program = Integer.parseInt(passData.Exceldata("Event Programs"));
	
	if (Event_Program>=5) 
	{			
		return events;
	}
	
	else if (Event_Program<5 && Event_Program>0 ) {
	
		return AccountName +"\nhas\n"+Exceldata("Event Programs")+"\nEvent campaigns in Marketo.\n" +events ;
	
	}
	else  {
	}
	
		return No_events  ;
				
} 


//Method of Nurture Campaigns 
public static String Nurture_Campaigns() throws IOException {	
	
	int Nurture_Campaigns = Integer.parseInt(passData.Exceldata("Nurture campaigns"));
	
	if (Nurture_Campaigns>=5) 
	{			
		return Exceldata("Nurture campaigns")+Nurture;
	}
	
	else if  (Nurture_Campaigns<5 && Nurture_Campaigns>0) {
	
		return Exceldata("Nurture campaigns")+Nurture ;
	
	}
	else  {
	}
	
		return No_Nurture  ;
				
}

//Segment Data 

public static String Segmentation() throws IOException {	
	
	int Segmentation = Integer.parseInt(passData.Exceldata("Segment Data"));
	
	if (Segmentation>=5) 
	{			
		return AccountName+"\nhas\n"+Exceldata("Segment Data")+Nurture;
	}
	
	else  {
	}
	
		return segment  ;
				
}



//Program Library 
public static String Library() throws IOException {	
	
	int Library = Integer.parseInt(passData.Exceldata("Library"));
	
	if (Library>=1) 
	{			
		return Program_library ;
	}
	
	else  {
	}
	
		return no_library ;
				
}

//Integrations

public static String Integrations() throws IOException {	
	
	int Integrations = Integer.parseInt(passData.Exceldata("Integrations"));
	
	if (Integrations>=1) 
	{			
		return passData.FetchScreenshot("Change Score") ;
	}
	
	else  {
	}
	
		return No_Integrations ;
				
}


	
		
}
	
	
	
	
	

