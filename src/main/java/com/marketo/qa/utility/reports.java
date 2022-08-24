package com.marketo.qa.utility;

import static org.testng.Assert.assertTrue;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.SimpleDateFormat;
import java.util.Date;


import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.BreakType;
import org.apache.poi.xwpf.usermodel.Document;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.UnderlinePatterns;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;



public class reports {
	

	

	static String word = System.getProperty("user.home") + "\\Desktop\\Reports\\";
	static String fileName = new SimpleDateFormat("dd_MM_yy_HH_mm").format(new Date());

	
	
	public static void docs() throws Exception {
		Path outputDirectory = Paths.get(word);
		if (!Files.exists(outputDirectory)) {
			assertTrue(new File(String.valueOf(outputDirectory)).mkdirs(), "Unable to create output directory");

		}

		
		
		
		XWPFDocument document = new XWPFDocument();
		
		XWPFParagraph Statsparagraph = document.createParagraph();
		XWPFRun run = Statsparagraph.createRun();
		run.setBold(true);
		run.setText("Stats");

		Statsparagraph = document.createParagraph();
		Statsparagraph.setSpacingAfter(0);
		run = Statsparagraph.createRun();
		stylingDoc.setNoProof(run);
		stylingDoc.FontFamilySize(run);
		Statsparagraph.setNumID(stylingDoc.bullet(document));
		run.setText("They have " + passData.Exceldata("All Campaigns")+ " campaigns");
		
		
		Statsparagraph = document.createParagraph();
		Statsparagraph.setSpacingAfter(0);
		run = Statsparagraph.createRun();
		stylingDoc.FontFamilySize(run);
		Statsparagraph.setNumID(stylingDoc.bullet(document));
		
		run.setText("They have " + passData.Exceldata("Active Campaigns") + " active campaigns");
		
		Statsparagraph = document.createParagraph();
		run = Statsparagraph.createRun();
		Statsparagraph.setSpacingAfter(0);
		stylingDoc.FontFamilySize(run);
		Statsparagraph.setNumID(stylingDoc.bullet(document));
		run.setText("They have " + passData.Exceldata("All Triggered Campaigns") + " triggered campaigns");
		
		Statsparagraph = document.createParagraph();
		run = Statsparagraph.createRun();
		Statsparagraph.setSpacingAfter(0);
		stylingDoc.FontFamilySize(run);
		Statsparagraph.setNumID(stylingDoc.bullet(document));
		run.setText("They have " + passData.Exceldata("Active Triggered Campaigns") + " Active triggered campaigns");
		
		Statsparagraph = document.createParagraph();
		run = Statsparagraph.createRun();
		Statsparagraph.setSpacingAfter(0);
		stylingDoc.FontFamilySize(run);
		Statsparagraph.setNumID(stylingDoc.bullet(document));
		run.setText("They have " + passData.Exceldata("Batch Campaigns - Repeating Schedule") + " re-occurring batch campaigns");
		
		Statsparagraph = document.createParagraph();
		run = Statsparagraph.createRun();
		stylingDoc.FontFamilySize(run);
		Statsparagraph.setSpacingAfter(0);
		Statsparagraph.setNumID(stylingDoc.bullet(document));
		run.setText("They have " + passData.Exceldata("All Batch Campaigns") + " batch campaigns");
		
		Statsparagraph = document.createParagraph();
		run = Statsparagraph.createRun();
		stylingDoc.FontFamilySize(run);
		Statsparagraph.setSpacingAfter(0);
		Statsparagraph.setNumID(stylingDoc.bullet(document));
		run.setText( passData.Exceldata("Landing Pages") + " landing pages");
		
		Statsparagraph = document.createParagraph();
		run = Statsparagraph.createRun();
		stylingDoc.FontFamilySize(run);
		Statsparagraph.setSpacingAfter(0);
		Statsparagraph.setNumID(stylingDoc.bullet(document));
		run.setText(passData.Exceldata("Forms")+" forms");
		
		Statsparagraph = document.createParagraph();
		run = Statsparagraph.createRun();
		stylingDoc.FontFamilySize(run);
		Statsparagraph.setSpacingAfter(0);
		Statsparagraph.setNumID(stylingDoc.bullet(document));
		run.setText(passData.Exceldata("Emails") + " emails");
		
		Statsparagraph = document.createParagraph();
		run = Statsparagraph.createRun();
		stylingDoc.FontFamilySize(run);
		Statsparagraph.setSpacingAfter(0);
		Statsparagraph.setNumID(stylingDoc.bullet(document));
		run.setText(passData.GenricMethod(passData.Exceldata("Snippets")));
		
		Statsparagraph = document.createParagraph();
		run = Statsparagraph.createRun();
		stylingDoc.FontFamilySize(run);
		Statsparagraph.setSpacingAfter(0);
		Statsparagraph.setNumID(stylingDoc.bullet(document));
		run.setText(passData.Exceldata("Images and Files") + " uploaded files");
		
		Statsparagraph = document.createParagraph();
		run = Statsparagraph.createRun();
		stylingDoc.FontFamilySize(run);
		Statsparagraph.setSpacingAfter(0);
		Statsparagraph.setNumID(stylingDoc.bullet(document));
		run.setText(passData.Exceldata("Leads") + "\nLeads");
		
		Statsparagraph = document.createParagraph();
		run = Statsparagraph.createRun();
		stylingDoc.FontFamilySize(run);
		Statsparagraph.setNumID(stylingDoc.bullet(document));
		run.setText(passData.Exceldata("Tags") + " programs");
		
		
		
		XWPFParagraph paragraphStats = document.createParagraph();
		XWPFRun r1 = paragraphStats.createRun();
		try {
			paragraphStats.setAlignment(ParagraphAlignment.LEFT);
			stylingDoc.setNoProof(r1);
			r1.addCarriageReturn();
			r1.addCarriageReturn();
			r1.setBold(false);
			
			stylingDoc.FontFamilySize(r1);
			r1.setText(passData.Exceldata("Account Name")+"," +passData.Org_info);
			
			XWPFRun r2 = paragraphStats.createRun();
			
			r2.setText("https://docs.marketo.com/display/public/DOCS/Smart+Campaigns");
			stylingDoc.setNoProof(r2);
			paragraphStats.createHyperlinkRun("https://docs.marketo.com/display/public/DOCS/Smart+Campaigns");
			r2.setUnderline(UnderlinePatterns.SINGLE);
			 r2.setColor("3333cc");
			 stylingDoc.FontFamilySize(r2);
			
			XWPFParagraph imgPara = document.createParagraph();
			XWPFRun img = imgPara.createRun();
			img.addPicture(new FileInputStream(passData.FetchScreenshot("Tags")), Document.PICTURE_TYPE_PNG, passData.FetchScreenshot("Tags"),
					Units.toEMU(380), Units.toEMU(150));
			img.addCarriageReturn();
			
			
			int Snippets = Integer.parseInt(passData.Exceldata("Snippets"));
			if (Snippets<5)
			{
				XWPFParagraph paragraphModel = document.createParagraph();
				XWPFRun model = paragraphModel.createRun();
				model.setBold(true);
				model.setText("Snippets");
				XWPFParagraph SnippetsData = document.createParagraph();
				XWPFRun SnippetsRun = SnippetsData.createRun();
				stylingDoc.setNoProof(SnippetsRun);
				stylingDoc.FontFamilySize(SnippetsRun);
				SnippetsRun.setText(passData.Exceldata("Account Name")+passData.No_Snippets);
				
				XWPFParagraph paragraphModel2 = document.createParagraph();
				XWPFRun modellink = paragraphModel2.createRun();
				stylingDoc.setNoProof(modellink);
				modellink.setText("\nhttps://experienceleague.adobe.com/docs/marketo/using/product-docs/personalization/segmentation-and-snippets/snippets/create-a-snippet.html?lang=en");
				paragraphModel.createHyperlinkRun("https://experienceleague.adobe.com/docs/marketo/using/product-docs/personalization/segmentation-and-snippets/snippets/create-a-snippet.html?lang=en");
				modellink.setUnderline(UnderlinePatterns.SINGLE);
				modellink.setColor("3333cc");
				stylingDoc.FontFamilySize(modellink);
				
			}
			
			
			//Models report part, Note : we have to added conditional based formatting 
			
			XWPFParagraph paragraphModel = document.createParagraph();
			XWPFRun model = paragraphModel.createRun();
			model.setBold(true);
			model.addBreak(BreakType.PAGE);
			model.addCarriageReturn();
			model.addCarriageReturn();
			model.setText("Models");
			
			int Models = Integer.parseInt(passData.Exceldata("Models"));;
			if (Models>0) 
					{
				XWPFParagraph paragraphModelData = document.createParagraph();
				XWPFRun modelData = paragraphModelData.createRun();
				stylingDoc.setNoProof(modelData);
				stylingDoc.FontFamilySize(modelData);
				
				
				
				modelData.setText(passData.Exceldata("Account Name") +",\nHas previously built\n"+passData.Exceldata("Models")+"\n models in Marketo. Revenue cycle models take marketing"
						+ " to the next level. They model all the stages of your entire revenue funnel—from when you first interact with a lead "
						+ "all the way until the lead is a won customer.");
				
				modelData = paragraphModelData.createRun();
				modelData.addBreak();
				stylingDoc.setNoProof(modelData);
				stylingDoc.FontFamilySize(modelData);
				modelData.setText(passData.Exceldata("Account Name")+passData.models2);
				
				
				
				XWPFParagraph model_img = document.createParagraph();
				XWPFRun img1 = model_img.createRun();
				img1.addPicture(new FileInputStream(passData.FetchScreenshot("Models")), Document.PICTURE_TYPE_PNG, passData.FetchScreenshot("Models"),
						Units.toEMU(230), Units.toEMU(150));
							
			}
			else
			{
				XWPFParagraph paragraphModelData = document.createParagraph();
				XWPFRun modelTest = paragraphModelData.createRun();
				stylingDoc.FontFamilySize(modelTest);
				stylingDoc.setNoProof(modelTest);
				modelTest.setText(passData.Exceldata("Account Name") + passData.No_models);
				
			}
			

			//adding lead scoring data into report 
			XWPFParagraph LeadParagraph = document.createParagraph();
			
			XWPFRun LeadRun = LeadParagraph.createRun();
			LeadRun.addCarriageReturn();
			LeadRun.addCarriageReturn();
			LeadRun.setBold(true);
			LeadRun.setText("Lead Scoring");
			
			
			
			// adding lead scoring bullet point data into report
			XWPFParagraph LeadData = document.createParagraph();
			
			XWPFRun LeadDataRun = LeadData.createRun();
			stylingDoc.FontFamilySize(LeadDataRun);
			stylingDoc.setNoProof(LeadDataRun);
			LeadData.setSpacingAfter(0);
			LeadData.setNumID(stylingDoc.bullet(document));
			LeadDataRun.setText(passData.LeadScoring(passData.Exceldata("Change Score")));
			
			int ChangeScore = Integer.parseInt(passData.Exceldata("Change Score"));
			if (ChangeScore>1) 	
			{
			XWPFParagraph LeadData1 = document.createParagraph();
			XWPFRun runData = LeadData1.createRun();
			stylingDoc.setNoProof(runData);
			LeadData1.setSpacingAfter(0);
			stylingDoc.FontFamilySize(runData);
			LeadData1.setNumID(stylingDoc.bullet(document));
			runData.setText(passData.Exceldata("Account Name") +","+"\nis executing multiple score changes with single campaigns.");
			}
			
			XWPFParagraph LeadData2 = document.createParagraph();
			XWPFRun runData2 = LeadData2.createRun();
			stylingDoc.FontFamilySize(runData2);
			LeadData2.setNumID(stylingDoc.bullet(document));
			stylingDoc.setNoProof(runData2);
			LeadData2.setSpacingAfter(0);
			runData2.setText(passData.MyTokens());
			
			XWPFParagraph LeadData3 = document.createParagraph();
			XWPFRun LeadDatarun3 = LeadData3.createRun();
			stylingDoc.FontFamilySize(LeadDatarun3);
			stylingDoc.setNoProof(LeadDatarun3);
			
			LeadData3.setNumID(stylingDoc.bullet(document));
			LeadDatarun3.setText(passData.BatchCampaigns());
		
			
			
			XWPFParagraph LeadData4 = document.createParagraph();
			XWPFRun LeadDataRun4 = LeadData4.createRun();
			
			stylingDoc.FontFamilySize(LeadDataRun4);
			stylingDoc.setNoProof(LeadDataRun4);
			LeadDataRun4.setText(passData.lead_scoring);
			
			
			//Interesting moment report part 
			if(passData.Exceldata("Interesting Moment Subscription").equalsIgnoreCase("true"))
			{
			
				XWPFParagraph Interesting = document.createParagraph();
				XWPFRun InterestingRun = Interesting.createRun();
				InterestingRun.setBold(true);
				InterestingRun.addCarriageReturn();
				InterestingRun.addCarriageReturn();
				InterestingRun.setText("Interesting Moment");
				int Interesting_Moment= Integer.parseInt(passData.Exceldata("Interesting Moment"));
				
				if (Interesting_Moment>0)
				{
				
				XWPFParagraph InterestingData = document.createParagraph();
				XWPFRun InterestingDatarun = InterestingData.createRun();
				
				stylingDoc.setNoProof(InterestingDatarun);
				stylingDoc.FontFamilySize(InterestingDatarun);
				InterestingDatarun.setText("The following Interesting Moments have been defined to support\n"+passData.intresting_moment);
				
				InterestingDatarun = InterestingData.createRun();
				InterestingDatarun.addPicture(new FileInputStream(passData.FetchScreenshot("Interesting Moment")), Document.PICTURE_TYPE_PNG, passData.FetchScreenshot("Interesting Moment"),
						Units.toEMU(470), Units.toEMU(80));
				
				 InterestingData = document.createParagraph();
				 InterestingDatarun = InterestingData.createRun();
			
				 stylingDoc.setNoProof(InterestingDatarun);
				 stylingDoc.FontFamilySize(InterestingDatarun);
				 InterestingDatarun.setText(passData.intresting_moment_below);
	
				}
				else {
					
					XWPFParagraph InterestingData = document.createParagraph();
					XWPFRun InterestingDatarun = InterestingData.createRun();
					stylingDoc.FontFamilySize(InterestingDatarun);
					stylingDoc.setNoProof(InterestingDatarun);
					InterestingDatarun.setText(passData.Exceldata("Account Name")+passData.intresting_moment_else);
					
					
				}
			}
			
			
		//adding Data management 	
			
			XWPFParagraph DataManagement = document.createParagraph();
			XWPFRun DataManagementRun = DataManagement.createRun();
			DataManagementRun.setBold(true);
			DataManagementRun.addCarriageReturn();
			DataManagementRun.addCarriageReturn();
			DataManagementRun.setText("Data Management");
			
			
			XWPFParagraph DataManagementData = document.createParagraph();
			XWPFRun DataManagementDatarun = DataManagementData.createRun();
			stylingDoc.FontFamilySize(DataManagementDatarun);
			stylingDoc.setNoProof(DataManagementDatarun);
			DataManagementDatarun.setText(passData.DataManagement());	

			
		//adding Events data 
			
			XWPFParagraph Events = document.createParagraph();
			XWPFRun	Eventsrun = Events.createRun();
			Eventsrun.setBold(true);
			Eventsrun.addCarriageReturn();
			
			Eventsrun.setText("Events");
	
			XWPFParagraph EventsData = document.createParagraph();
			
			XWPFRun EventsDatarun = EventsData.createRun();
			stylingDoc.FontFamilySize(EventsDatarun);
			stylingDoc.setNoProof(EventsDatarun);
			EventsDatarun.setText(passData.Events());
			
			
			
		
		//adding Nurture data 
			
			
			XWPFParagraph NurtureData = document.createParagraph();
			XWPFRun NurtureDataRun = NurtureData.createRun();
			NurtureDataRun.setBold(true);
			NurtureDataRun.addCarriageReturn();
			NurtureDataRun.setText("Nurture");
			NurtureDataRun.addCarriageReturn();
			
			int NurtureDataint= Integer.parseInt(passData.Exceldata("Nurture campaigns"));
			
			if (NurtureDataint>0)
				
			{
			NurtureDataRun = NurtureData.createRun();
			stylingDoc.FontFamilySize(NurtureDataRun);
			stylingDoc.setNoProof(NurtureDataRun);
			NurtureDataRun.setText(passData.Exceldata("Account Name")+"\nhas\n"+passData.Exceldata("Nurture campaigns")+passData.Nurture);
			
			}
			else
			{
				NurtureDataRun = NurtureData.createRun();
				stylingDoc.FontFamilySize(NurtureDataRun);
				stylingDoc.setNoProof(NurtureDataRun);
				NurtureDataRun.setText(passData.Exceldata("Account Name")+passData.No_Nurture);
				
				XWPFRun NurtureDataRunlink = NurtureData.createRun();
				stylingDoc.FontFamilySize(NurtureDataRunlink);
			//	NurtureData.createHyperlinkRun("https://experienceleague.adobe.com/docs/marketo/using/product-docs/email-marketing/drip-nurturing/creating-an-engagement-program/create-an-engagement-program.html?lang=en\r\n");
				NurtureDataRunlink.setUnderline(UnderlinePatterns.SINGLE);
				NurtureDataRunlink.setColor("3333cc");
				NurtureDataRunlink.setText("https://experienceleague.adobe.com/docs/marketo/using/product-docs/email-marketing/drip-nurturing/creating-an-engagement-program/create-an-engagement-program.html?lang=en");
				
				
			}
			
			
			XWPFParagraph NurtureData1 = document.createParagraph();
			XWPFRun	NurtureDatarun1 = NurtureData1.createRun();
			stylingDoc.FontFamilySize(NurtureDatarun1);
			NurtureData1.setNumID(stylingDoc.bullet(document));
			NurtureData1.setSpacingAfter(0);
			NurtureDatarun1.setText("Intelligently and automatically deliver content to a target audience.");
			
			
			XWPFParagraph NurtureData2 = document.createParagraph();
			XWPFRun NurtureDatarun2 = NurtureData2.createRun();
			stylingDoc.FontFamilySize(NurtureDatarun2);
			NurtureData2.setSpacingAfter(0);
			NurtureData2.setNumID(stylingDoc.bullet(document));
			NurtureDatarun2.setText("Easily build dialogue with prospects and customers while preventing customers who have already received content from receiving the same content again.");
			
			XWPFParagraph NurtureData3 = document.createParagraph();
			XWPFRun	NurtureDatarun3 = NurtureData3.createRun();
			stylingDoc.FontFamilySize(NurtureDatarun3);
			NurtureData3.setSpacingAfter(0);
			NurtureData3.setNumID(stylingDoc.bullet(document));
			NurtureDatarun3.setText("Add new content and entire programs to nurture streams.");
			
			XWPFParagraph NurtureData4 = document.createParagraph();
			XWPFRun NurtureDatarun4 = NurtureData4.createRun();
			stylingDoc.FontFamilySize(NurtureDatarun4);
			NurtureData4.setSpacingAfter(0);
			NurtureData4.setNumID(stylingDoc.bullet(document));
			NurtureDatarun4.setText("Edit the availability of content.");
			
			XWPFParagraph NurtureData5 = document.createParagraph();
			XWPFRun	NurtureDatarun5 = NurtureData5.createRun();
			stylingDoc.FontFamilySize(NurtureDatarun5);
			NurtureData5.setNumID(stylingDoc.bullet(document));
			NurtureDatarun5.setText("Understand content performance based on engagement with each piece of content.");
			NurtureDatarun5.addCarriageReturn();
			
			
			//Segment Data printing 
			XWPFParagraph Segment = document.createParagraph();
			XWPFRun SegmentHeading = Segment.createRun();
			SegmentHeading.setBold(true);
			SegmentHeading.setText("Segmentation");
			
			int Segment_Data= Integer.parseInt(passData.Exceldata("Segmentations"));
			if (Segment_Data>0)
			{
			
			XWPFParagraph SegmentData = document.createParagraph();
			XWPFRun SegmentRun = SegmentData.createRun();
			stylingDoc.FontFamilySize(SegmentRun);
			stylingDoc.setNoProof(SegmentRun);
			SegmentRun.setText(passData.Exceldata("Account Name")+passData.segment);
			
			
			SegmentRun = SegmentData.createRun();
			SegmentRun.addBreak();
			SegmentRun.addBreak();
			SegmentRun.addPicture(new FileInputStream(passData.FetchScreenshot("Segmentations")), Document.PICTURE_TYPE_PNG, passData.FetchScreenshot("Segmentations"),
					Units.toEMU(230), Units.toEMU(50));

			}
			else {
				
				XWPFParagraph InterestingData = document.createParagraph();
				XWPFRun InterestingDatarun = InterestingData.createRun();
				stylingDoc.FontFamilySize(InterestingDatarun);
				stylingDoc.setNoProof(InterestingDatarun);
				InterestingDatarun.setText(passData.Exceldata("Account Name")+passData.no_segment);
				InterestingDatarun.addCarriageReturn();
				
				
			}
			
			//program library section
			
			XWPFParagraph program = document.createParagraph();
			XWPFRun programRun = program.createRun();
			programRun.addCarriageReturn();
			programRun.addCarriageReturn();
			programRun.setBold(true);
			programRun.setText("Program Library");
			
			
			
			XWPFParagraph ProgramData = document.createParagraph();
			XWPFRun ProgramDataRun = ProgramData.createRun();
			stylingDoc.FontFamilySize(ProgramDataRun);
			ProgramDataRun.addCarriageReturn();
			stylingDoc.setNoProof(ProgramDataRun);
			ProgramDataRun.setText(passData.Library());
	
			
			
			
			//Integration Data Section
			
			XWPFParagraph intigration = document.createParagraph();
			XWPFRun intigrationRunHeading = intigration.createRun();
			intigrationRunHeading.addCarriageReturn();
			intigrationRunHeading.addCarriageReturn();
			intigrationRunHeading.setBold(true);
			intigrationRunHeading.setText("Integrations");
			
			int intigration_Data= Integer.parseInt(passData.Exceldata("Integration"));
			if (intigration_Data>0)
			{
			
			XWPFParagraph intigrationRun = document.createParagraph();
			XWPFRun intigrationData = intigrationRun.createRun();
			stylingDoc.FontFamilySize(intigrationData);
			intigrationData.setText("The following integrations have been installed:");
	
			intigrationRun = document.createParagraph();
			XWPFRun intigrationImg = intigrationRun.createRun();
			intigrationImg.addPicture(new FileInputStream(passData.FetchScreenshot("Integration")), Document.PICTURE_TYPE_PNG, passData.FetchScreenshot("Integration"),
					Units.toEMU(470), Units.toEMU(200));

			}
			else {
				
				XWPFParagraph intigrationRun = document.createParagraph();
				XWPFRun intigrationData = intigrationRun.createRun();
				stylingDoc.FontFamilySize(intigrationData);
				stylingDoc.setNoProof(intigrationData);
				intigrationData.setText(passData.Exceldata("Account Name")+passData.No_Integrations);
				
				
			}
		
			
			
			//Web Personalize Data Section
			if(passData.Exceldata("Web Personalize")=="true")
			{
			XWPFParagraph Web_Personalize = document.createParagraph();
			XWPFRun Web_PersonalizeRun = Web_Personalize.createRun();
			Web_PersonalizeRun.addCarriageReturn();
			Web_PersonalizeRun.addCarriageReturn();
			Web_PersonalizeRun.setBold(true);
			Web_PersonalizeRun.setText("Web Personalize");
			
			
			
			XWPFParagraph PersonalizeDoc = document.createParagraph();
			XWPFRun PersonalizeRun = PersonalizeDoc.createRun();
			stylingDoc.FontFamilySize(PersonalizeRun);
			stylingDoc.setNoProof(PersonalizeRun);
			PersonalizeRun.setText(passData.web_personalize);
			
			
			PersonalizeDoc = document.createParagraph();
			XWPFRun PersonalizeImg = PersonalizeDoc.createRun();
			stylingDoc.setNoProof(PersonalizeImg);
			stylingDoc.FontFamilySize(PersonalizeImg);
			PersonalizeImg.setText("Top Campaigns: The top performing campaigns during the selected time period, ordered by number of clicks.");
			PersonalizeImg.addBreak();
			XWPFRun PersonalizeImg1 = PersonalizeDoc.createRun();
			PersonalizeImg1.addPicture(new FileInputStream(passData.FetchScreenshot("Change Score")), Document.PICTURE_TYPE_PNG, passData.FetchScreenshot("Change Score"),
					Units.toEMU(470), Units.toEMU(80));
			
			}
			
			//Target Account Management
			if(passData.Exceldata("Target Account Management")=="true")
			{
			XWPFParagraph Target_Account = document.createParagraph();
			XWPFRun Target_AccountRun = Target_Account.createRun();
			
			Target_AccountRun.addCarriageReturn();
			Target_AccountRun.addCarriageReturn();
			Target_AccountRun.setBold(true);
			Target_AccountRun.setText("Target Account Management");
			
			
			
			XWPFParagraph Account_Management = document.createParagraph();
			XWPFRun Account_ManagementRun = Account_Management.createRun();
			stylingDoc.FontFamilySize(Account_ManagementRun);
			stylingDoc.setNoProof(Account_ManagementRun);
			Account_ManagementRun.setText(passData.Account_Management);
			Account_ManagementRun.addCarriageReturn();
			Account_ManagementRun.setText(passData.Account_Management2);
			
			
			XWPFParagraph Overview = document.createParagraph();
			XWPFRun OverviewRun = Overview.createRun();
			OverviewRun.addCarriageReturn();
			OverviewRun.addCarriageReturn();
			stylingDoc.FontFamilySize(OverviewRun);
			OverviewRun.setBold(true);
			OverviewRun.setText("Overview");
			OverviewRun.addBreak();
			XWPFRun OverviewRunImg1 = Account_Management.createRun();
			OverviewRunImg1.addPicture(new FileInputStream(passData.FetchScreenshot("Change Score")), Document.PICTURE_TYPE_PNG, passData.FetchScreenshot("Change Score"),
					Units.toEMU(470), Units.toEMU(80));
			
			

			}
			
			
					
		} catch (Exception e) {

		}
		stylingDoc.HeaderFooter(document);// Styling of Header and footer
		
		stylingDoc.bold(document);
		FileOutputStream out = new FileOutputStream(word +passData.Exceldata("Account Name")+fileName+".docx");

		document.write(out);
		out.close();

	}

	
		
	}


