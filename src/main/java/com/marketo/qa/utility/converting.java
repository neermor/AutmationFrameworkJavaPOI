package com.marketo.qa.utility;

import java.io.IOException;

public class converting {

	static int AllBatchCampaigns = 0;
	static int AllCampaigns = 0;
	static int ActiveChampaigns = 0;
	static int AllTriggeredCampaigns = 0;
	static int ActiveTriggeredCampaigns = 0;
	static int BatchCampaignsRepeatingSchedule = 0;
	static int LandingPages = 0;
	static int Forms = 0;
	static int Emails = 0;
	static int Snippets = 0;
	static int ImagesandFiles = 0;
	static int AllPeople = 0;
	static int MarketableLeads = 0;
	static int PossibleDuplicates = 0;
	static int UnsubscribedPeople = 0;
	static int MarketingSuspended = 0;
	static int tags = 0;
	static int Models = 0;
	static int ScheduledBatchCampaigns = 0;
	static int Integration=0;
	static int DataManagement=0;
	

	public static void localtest() throws NumberFormatException, IOException {
		try {
		AllCampaigns = Integer.parseInt(passData.Exceldata("All Campaigns"));

		AllBatchCampaigns = Integer.parseInt(passData.Exceldata("All Batch Campaigns"));

		ActiveChampaigns = Integer.parseInt(passData.Exceldata("Active Campaigns"));

		AllTriggeredCampaigns = Integer.parseInt(passData.Exceldata("All Triggered Campaigns"));

		ActiveTriggeredCampaigns = Integer.parseInt(passData.Exceldata("Active Triggered Campaigns"));

		BatchCampaignsRepeatingSchedule = Integer.parseInt(passData.Exceldata("Batch Campaigns - Repeating Schedule"));

		LandingPages = Integer.parseInt(passData.Exceldata("Landing Pages"));

		Forms = Integer.parseInt(passData.Exceldata("Forms"));

		Emails = Integer.parseInt(passData.Exceldata("Emails"));

		Snippets = Integer.parseInt(passData.Exceldata("Snippets"));

		ImagesandFiles = Integer.parseInt(passData.Exceldata("Images and Files"));

		AllPeople = Integer.parseInt(passData.Exceldata("All People"));

		MarketableLeads = Integer.parseInt(passData.Exceldata("Marketable Leads"));

		PossibleDuplicates = Integer.parseInt(passData.Exceldata("Possible Duplicates"));

		UnsubscribedPeople = Integer.parseInt(passData.Exceldata("Unsubscribed People"));

		MarketingSuspended = Integer.parseInt(passData.Exceldata("Marketing Suspended"));

		tags = Integer.parseInt(passData.Exceldata("Tags"));

		Models = Integer.parseInt(passData.Exceldata("Models"));
		
		Integration= Integer.parseInt(passData.Exceldata("Integration"));
		DataManagement = Integer.parseInt(passData.Exceldata("Change Data Value"));
		
		ScheduledBatchCampaigns = Integer.parseInt(passData.Exceldata("Batch Campaigns - Repeating Schedule"));
		}
		catch (Exception e) {
			// TODO: handle exception
		}

	}
}
