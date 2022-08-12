package com.marketo.qa.utility;

import static org.testng.Assert.assertTrue;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.BreakType;
import org.apache.poi.xwpf.usermodel.Document;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.xmlbeans.XmlException;





public class reports {
	

	

	static String word = System.getProperty("user.home") + "\\Desktop\\Reports\\";
	static String fileName = new SimpleDateFormat("dd_MM_yy_HH_mm").format(new Date());
	static String screenshotsPathWordDoc = word + "Report_"+fileName+".docx";


	public static void docs() throws IOException, XmlException {
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
		run = Statsparagraph.createRun();
		stylingDoc.setNoProof(run);
		run.setFontFamily("Calibri Light (Headings)");
		run.setFontSize(10);
		Statsparagraph.setNumID(stylingDoc.bullet(document));
		run.setText("They have " + passData.Exceldata("All Campaigns")+ " campaigns");
		
		
		Statsparagraph = document.createParagraph();
		run = Statsparagraph.createRun();
		run.setFontFamily("Calibri Light (Headings)");
		run.setFontSize(10);
		Statsparagraph.setNumID(stylingDoc.bullet(document));
		run.setText("They have " + passData.Exceldata("Active Campaigns") + " active campaigns");
		
		Statsparagraph = document.createParagraph();
		run = Statsparagraph.createRun();
		run.setFontFamily("Calibri Light (Headings)");
		run.setFontSize(10);
		Statsparagraph.setNumID(stylingDoc.bullet(document));
		run.setText("They have " + passData.Exceldata("All Triggered Campaigns") + " triggered campaigns");
		
		Statsparagraph = document.createParagraph();
		run = Statsparagraph.createRun();
		run.setFontFamily("Calibri Light (Headings)");
		run.setFontSize(10);
		Statsparagraph.setNumID(stylingDoc.bullet(document));
		run.setText("They have " + passData.Exceldata("Active Triggered Campaigns") + " Active triggered campaigns");
		
		Statsparagraph = document.createParagraph();
		run = Statsparagraph.createRun();
		run.setFontFamily("Calibri Light (Headings)");
		run.setFontSize(10);
		Statsparagraph.setNumID(stylingDoc.bullet(document));
		run.setText("They have " + passData.Exceldata("Active Triggered Campaigns") + " re-occurring batch campaigns");
		
		Statsparagraph = document.createParagraph();
		run = Statsparagraph.createRun();
		run.setFontFamily("Calibri Light (Headings)");
		run.setFontSize(10);
		Statsparagraph.setNumID(stylingDoc.bullet(document));
		run.setText("They have " + passData.Exceldata("All Batch Campaigns") + " batch campaigns");
		
		Statsparagraph = document.createParagraph();
		run = Statsparagraph.createRun();
		run.setFontFamily("Calibri Light (Headings)");
		run.setFontSize(10);
		Statsparagraph.setNumID(stylingDoc.bullet(document));
		run.setText( passData.Exceldata("Landing Pages") + " landing pages");
		
		Statsparagraph = document.createParagraph();
		run = Statsparagraph.createRun();
		run.setFontFamily("Calibri Light (Headings)");
		run.setFontSize(10);
		Statsparagraph.setNumID(stylingDoc.bullet(document));
		run.setText(passData.Exceldata("Forms")+" forms");
		
		Statsparagraph = document.createParagraph();
		run = Statsparagraph.createRun();
		run.setFontFamily("Calibri Light (Headings)");
		run.setFontSize(10);
		Statsparagraph.setNumID(stylingDoc.bullet(document));
		run.setText(passData.Exceldata("Emails") + " emails");
		
		Statsparagraph = document.createParagraph();
		run = Statsparagraph.createRun();
		run.setFontFamily("Calibri Light (Headings)");
		run.setFontSize(10);
		Statsparagraph.setNumID(stylingDoc.bullet(document));
		run.setText(passData.GenricMethod(passData.Exceldata("Snippets")));
		
		Statsparagraph = document.createParagraph();
		run = Statsparagraph.createRun();
		run.setFontFamily("Calibri Light (Headings)");
		run.setFontSize(10);
		Statsparagraph.setNumID(stylingDoc.bullet(document));
		run.setText(passData.Exceldata("Images and Files") + " uploaded files");
		
		Statsparagraph = document.createParagraph();
		run = Statsparagraph.createRun();
		run.setFontFamily("Calibri Light (Headings)");
		run.setFontSize(10);
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
			
			r1.setFontFamily("Calibri Light (Headings)");
			r1.setFontSize(10);
			r1.setText(passData.AccountName +"," +passData.Org_info);
			
			XWPFParagraph imgPara = document.createParagraph();
			XWPFRun img = imgPara.createRun();
			img.addPicture(new FileInputStream(passData.FetchScreenshot("Tags")), Document.PICTURE_TYPE_PNG, passData.FetchScreenshot("Tags"),
					Units.toEMU(380), Units.toEMU(150));
			
			img.addBreak(BreakType.PAGE);
			
			
			//Models report part, Note : we have to added conditional based formatting 
			
			XWPFParagraph paragraphModel = document.createParagraph();
			XWPFRun model = paragraphModel.createRun();
			model.setBold(true);
			model.setText("Models");
			
			File f = new File(passData.FetchScreenshot( "Tags"));
			if (f.exists()) 
					{
				XWPFParagraph paragraphModelData = document.createParagraph();
				XWPFRun modelData = paragraphModelData.createRun();
				stylingDoc.setNoProof(modelData);
				modelData.setFontFamily("Calibri Light (Headings)");
				modelData.setFontSize(10);
				
				
				
				modelData.setText(passData.AccountName +",\nHas\n"+passData.Exceldata("Models")+"\nbuilt models in Marketo. Revenue cycle models take marketing"
						+ " to the next level. They model all the stages of your entire revenue funnel—from when you first interact with a lead "
						+ "all the way until the lead is a won customer.");
				
				modelData = paragraphModelData.createRun();
				modelData.setFontFamily("Calibri Light (Headings)");
				modelData.addBreak();
				modelData.setFontSize(10);
				stylingDoc.setNoProof(modelData);
				modelData.setText(passData.models2);
				
				
				
				XWPFParagraph model_img = document.createParagraph();
				XWPFRun img1 = model_img.createRun();
				img1.addPicture(new FileInputStream(passData.FetchScreenshot("Tags")), Document.PICTURE_TYPE_PNG, passData.FetchScreenshot("Tags"),
						Units.toEMU(380), Units.toEMU(150));
							
			}
			else
			{
				
				XWPFParagraph paragraphModelData = document.createParagraph();
				XWPFRun modelTest = paragraphModelData.createRun();
				modelTest.setFontFamily("Calibri Light (Headings)");
				modelTest.setFontSize(10);
				stylingDoc.setNoProof(modelTest);
				modelTest.setText(passData.No_models);
				modelTest.addCarriageReturn();
				modelTest.addCarriageReturn();  
			}
			
			
			
			
			
			//adding lead scoring data into report 
			XWPFParagraph LeadParagraph = document.createParagraph();
			
			XWPFRun LeadRun = LeadParagraph.createRun();
			LeadRun.setBold(true);
			LeadRun.setText("Lead Scoring");
			
			
			
			// adding lead scoring bullet point data into report
			XWPFParagraph LeadData = document.createParagraph();
			
			XWPFRun LeadDataRun = LeadData.createRun();
			LeadDataRun.setFontFamily("Calibri Light (Headings)");
			LeadDataRun.setFontSize(10);
			stylingDoc.setNoProof(LeadDataRun);
			LeadData.setNumID(stylingDoc.bullet(document));
			LeadDataRun.setText(passData.LeadScoring(passData.Exceldata("Leads")));
			
			XWPFParagraph LeadData1 = document.createParagraph();
			XWPFRun runData = LeadData1.createRun();
			stylingDoc.setNoProof(runData);
			runData.setFontFamily("Calibri Light (Headings)");
			runData.setFontSize(10);
			LeadData1.setNumID(stylingDoc.bullet(document));
			runData.setText(passData.AccountName +","+"\nis executing multiple score changes with single campaigns.");
			
			XWPFParagraph LeadData2 = document.createParagraph();
			XWPFRun runData2 = LeadData2.createRun();
			runData2.setFontFamily("Calibri Light (Headings)");
			runData2.setFontSize(10);
			LeadData2.setNumID(stylingDoc.bullet(document));
			stylingDoc.setNoProof(runData2);
			runData2.setText(passData.MyTokens());
			
			XWPFParagraph LeadData3 = document.createParagraph();
			XWPFRun LeadDatarun3 = LeadData3.createRun();
			LeadDatarun3.setFontFamily("Calibri Light (Headings)");
			LeadDatarun3.setFontSize(10);
			stylingDoc.setNoProof(LeadDatarun3);
			LeadData3.setNumID(stylingDoc.bullet(document));
			LeadDatarun3.setText(passData.BatchCampaigns());
		
			
			
			XWPFParagraph LeadData4 = document.createParagraph();
			XWPFRun LeadDataRun4 = LeadData4.createRun();
			LeadDataRun4.setFontFamily("Calibri Light (Headings)");
			LeadDataRun4.setFontSize(10);
			stylingDoc.setNoProof(LeadDataRun4);
			LeadDataRun4.setText(passData.lead_scoring);
			
			
			//Interesting moment report part 
			if(passData.Exceldata("Interesting Moment_subscription")=="true")
			{
			
				XWPFParagraph Interesting = document.createParagraph();
				XWPFRun InterestingRun = Interesting.createRun();
				InterestingRun.setBold(true);
				InterestingRun.setText("Interesting Moment");
				int Interesting_Moment= Integer.parseInt(passData.Exceldata("Interesting Moment"));
				
				if (Interesting_Moment>0)
				{
				
				XWPFParagraph InterestingData = document.createParagraph();
				XWPFRun InterestingDatarun = InterestingData.createRun();
				InterestingDatarun.setFontFamily("Calibri Light (Headings)");
				stylingDoc.setNoProof(InterestingDatarun);
				InterestingDatarun.setFontSize(10);
				InterestingDatarun.setText(passData.intresting_moment);
				
				InterestingDatarun = InterestingData.createRun();
				InterestingDatarun.addPicture(new FileInputStream(passData.FetchScreenshot("Interesting Moment")), Document.PICTURE_TYPE_PNG, passData.FetchScreenshot("Interesting Moment"),
						Units.toEMU(470), Units.toEMU(80));
				
				 InterestingData = document.createParagraph();
				 InterestingDatarun = InterestingData.createRun();
				 InterestingDatarun.setFontSize(10);
				 stylingDoc.setNoProof(InterestingDatarun);
				 InterestingDatarun.setFontFamily("Calibri Light (Headings)");
				 InterestingDatarun.setText(passData.intresting_moment_below);
	
				}
				else {
					
					XWPFParagraph InterestingData = document.createParagraph();
					XWPFRun InterestingDatarun = InterestingData.createRun();
					InterestingDatarun.setFontFamily("Calibri Light (Headings)");
					InterestingDatarun.setFontSize(10);
					stylingDoc.setNoProof(InterestingDatarun);
					InterestingDatarun.setText(passData.intresting_moment_else);
					InterestingDatarun.addCarriageReturn();
					
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
			DataManagementDatarun.setFontFamily("Calibri Light (Headings)");
			DataManagementDatarun.setFontSize(10);
			stylingDoc.setNoProof(DataManagementDatarun);
			DataManagementDatarun.setText(passData.Data);	

			
		//adding Events data 
			
			XWPFParagraph Events = document.createParagraph();
			XWPFRun	Eventsrun = Events.createRun();
			Eventsrun.setBold(true);
			Eventsrun.addCarriageReturn();
			
			Eventsrun.setText("Events");
	
			XWPFParagraph EventsData = document.createParagraph();
			XWPFRun EventsDatarun = EventsData.createRun();
			EventsDatarun.setFontFamily("Calibri Light (Headings)");
			EventsDatarun.setFontSize(10);
			stylingDoc.setNoProof(EventsDatarun);
			EventsDatarun.setText(passData.AccountName  + " has built numerous Event campaigns in Marketo.");
			
			
			EventsDatarun = EventsData.createRun();
			EventsDatarun.setFontFamily("Calibri Light (Headings)");
			EventsDatarun.setFontSize(10);
			stylingDoc.setNoProof(EventsDatarun);
			EventsDatarun.setText(passData.events);
			EventsDatarun.addBreak();
			
			
		
		//adding Nurture data 
			
			
			XWPFParagraph NurtureData = document.createParagraph();
			XWPFRun NurtureDataRun = NurtureData.createRun();
			NurtureDataRun.setBold(true);
			NurtureDataRun.setText("Nurture");
			NurtureDataRun.addBreak();
			
			NurtureDataRun = NurtureData.createRun();
			NurtureDataRun.setFontFamily("Calibri Light (Headings)");
			NurtureDataRun.setFontSize(10);
			stylingDoc.setNoProof(NurtureDataRun);
			NurtureDataRun.setText(passData.Nurture_Campaigns());
			
		
			XWPFParagraph NurtureData1 = document.createParagraph();
			XWPFRun	NurtureDatarun1 = NurtureData1.createRun();
			NurtureDatarun1.setFontFamily("Calibri Light (Headings)");
			NurtureDatarun1.setFontSize(10);
			NurtureData1.setNumID(stylingDoc.bullet(document));
			NurtureDatarun1.setText("Intelligently and automatically deliver content to a target audience.");
			
			XWPFParagraph NurtureData2 = document.createParagraph();
			XWPFRun NurtureDatarun2 = NurtureData2.createRun();
			NurtureDatarun2.setFontFamily("Calibri Light (Headings)");
			NurtureDatarun2.setFontSize(10);
			NurtureData2.setNumID(stylingDoc.bullet(document));
			NurtureDatarun2.setText("Easily build dialogue with prospects and customers while preventing customers who have already received content from receiving the same content again.");
			
			XWPFParagraph NurtureData3 = document.createParagraph();
			XWPFRun	NurtureDatarun3 = NurtureData3.createRun();
			NurtureDatarun3.setFontFamily("Calibri Light (Headings)");
			NurtureDatarun3.setFontSize(10);
			NurtureData3.setNumID(stylingDoc.bullet(document));
			NurtureDatarun3.setText("Add new content and entire programs to nurture streams.");
			
			XWPFParagraph NurtureData4 = document.createParagraph();
			XWPFRun NurtureDatarun4 = NurtureData4.createRun();
			NurtureDatarun4.setFontFamily("Calibri Light (Headings)");
			NurtureDatarun4.setFontSize(10);
			NurtureData4.setNumID(stylingDoc.bullet(document));
			NurtureDatarun4.setText("Edit the availability of content.");
			
			XWPFParagraph NurtureData5 = document.createParagraph();
			XWPFRun	NurtureDatarun5 = NurtureData5.createRun();
			NurtureDatarun5.setFontFamily("Calibri Light (Headings)");
			NurtureDatarun5.setFontSize(10);
			NurtureData5.setNumID(stylingDoc.bullet(document));
			NurtureDatarun5.setText("Understand content performance based on engagement with each piece of content.");
			NurtureDatarun5.addBreak(BreakType.PAGE);
			
			
			//Segment Data printing 
			XWPFParagraph Segment = document.createParagraph();
			XWPFRun SegmentHeading = Segment.createRun();
			SegmentHeading.setBold(true);
			SegmentHeading.setText("Segment");
			
			int Segment_Data= Integer.parseInt(passData.Exceldata("Segment Data"));
			if (Segment_Data>0)
			{
			
			XWPFParagraph SegmentData = document.createParagraph();
			XWPFRun SegmentRun = SegmentData.createRun();
			SegmentRun.setFontFamily("Calibri Light (Headings)");
			SegmentRun.setFontSize(10);
			stylingDoc.setNoProof(SegmentRun);
			SegmentRun.setText(passData.segment);
			
			SegmentRun = SegmentData.createRun();
			SegmentRun.addPicture(new FileInputStream(passData.FetchScreenshot("Change Score")), Document.PICTURE_TYPE_PNG, passData.FetchScreenshot("Change Score"),
					Units.toEMU(470), Units.toEMU(80));

			}
			else {
				
				XWPFParagraph InterestingData = document.createParagraph();
				XWPFRun InterestingDatarun = InterestingData.createRun();
				InterestingDatarun.setFontFamily("Calibri Light (Headings)");
				InterestingDatarun.setFontSize(10);
				stylingDoc.setNoProof(InterestingDatarun);
				InterestingDatarun.setText(passData.no_segment);
				InterestingDatarun.addCarriageReturn();
				
				
			}
			
			//Integration Data Section
			
			XWPFParagraph intigration = document.createParagraph();
			XWPFRun intigrationRunHeading = intigration.createRun();
			intigrationRunHeading.addCarriageReturn();
			intigrationRunHeading.addCarriageReturn();
			intigrationRunHeading.setBold(true);
			intigrationRunHeading.setText("Integrations");
			
			int intigration_Data= Integer.parseInt(passData.Exceldata("Integration Data"));
			if (intigration_Data>0)
			{
			
			XWPFParagraph intigrationRun = document.createParagraph();
			XWPFRun intigrationData = intigrationRun.createRun();
			intigrationData.setFontFamily("Calibri Light (Headings)");
			
			intigrationData.setFontSize(10);
			intigrationData.setText("The following integrations have been installed:");
	
			intigrationRun = document.createParagraph();
			XWPFRun intigrationImg = intigrationRun.createRun();
			intigrationImg.addPicture(new FileInputStream(passData.FetchScreenshot("Change Score")), Document.PICTURE_TYPE_PNG, passData.FetchScreenshot("Change Score"),
					Units.toEMU(470), Units.toEMU(80));

			}
			else {
				
				XWPFParagraph intigrationRun = document.createParagraph();
				XWPFRun intigrationData = intigrationRun.createRun();
				intigrationData.setFontFamily("Calibri Light (Headings)");
				intigrationData.setFontSize(10);
				stylingDoc.setNoProof(intigrationData);
				intigrationData.setText(passData.No_Integrations);
				intigrationData.addCarriageReturn();
				
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
			PersonalizeRun.setFontFamily("Calibri Light (Headings)");
			PersonalizeRun.setFontSize(10);
			stylingDoc.setNoProof(PersonalizeRun);
			PersonalizeRun.setText(passData.web_personalize);
			
			
			PersonalizeDoc = document.createParagraph();
			XWPFRun PersonalizeImg = PersonalizeDoc.createRun();
			stylingDoc.setNoProof(PersonalizeImg);
			PersonalizeRun.setFontSize(10);
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
			Account_ManagementRun.setFontFamily("Calibri Light (Headings)");
			Account_ManagementRun.setFontSize(10);
			stylingDoc.setNoProof(Account_ManagementRun);
			Account_ManagementRun.setText(passData.Account_Management);
			Account_ManagementRun.addCarriageReturn();
			Account_ManagementRun.setText(passData.Account_Management2);
			
			
			XWPFParagraph Overview = document.createParagraph();
			XWPFRun OverviewRun = Overview.createRun();
			OverviewRun.addCarriageReturn();
			OverviewRun.addCarriageReturn();
			OverviewRun.setFontFamily("Calibri Light (Headings)");
			OverviewRun.setFontSize(10);
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
		FileOutputStream out = new FileOutputStream(screenshotsPathWordDoc);

		document.write(out);
		out.close();

	}

	
		
	}


