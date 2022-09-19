package com.marketo.qa.utility;


import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import org.apache.logging.log4j.*;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.BreakType;
import org.apache.poi.xwpf.usermodel.Document;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.UnderlinePatterns;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.xmlbeans.XmlException;




public class docReports  {
	private static Logger logger =LogManager.getLogger(docReports.class.getName()); 

	static String word = System.getProperty("user.dir") + "/Reports/";
	static String fileName = new SimpleDateFormat("dd_MM_yy_HH_mm").format(new Date());
	
	
	
	public static void close(XWPFDocument document) throws InvalidFormatException, IOException
	{
		stylingDoc.HeaderFooter(document);// Styling of Header and footer

		stylingDoc.bold(document);
		FileOutputStream out = new FileOutputStream(word + passData.Exceldata("Account Name") + fileName + ".docx");

		document.write(out);
		out.close();
		logger.info("Doc file is Ready for " +passData.Exceldata("Account Name"));
	}

	public static void stats(XWPFDocument document) throws NumberFormatException, InvalidFormatException, FileNotFoundException, IOException, XmlException
	{
		XWPFParagraph Statsparagraph = document.createParagraph();
		XWPFRun run = Statsparagraph.createRun();
		Statsparagraph = document.createParagraph();
		run = Statsparagraph.createRun();
		stylingDoc.setNoProof(run);
		stylingDoc.FontFamilySize(run);
		logger.info("Start writing in Doc file");
		logger.info("Printing workspaces");
		try {
		run.setText(passData.Exceldata("Account Name") + "\nhas\n"
				+ passData.Exceldata("Total WorkSpace")+"\nworkspaces. Stats below are a sum of assets found across all workspaces.");
		}
		catch (Exception e) {
			// TODO: handle exception
			logger.info("Work Space counting is missing\n");
		}
		Statsparagraph = document.createParagraph();
		run = Statsparagraph.createRun();
		run.addCarriageReturn();
		run.setBold(true);
		logger.info("Printing Stats Section.....");
		run.setText("Stats");

		Statsparagraph = document.createParagraph();
		Statsparagraph.setSpacingAfter(0);
		run = Statsparagraph.createRun();
		stylingDoc.setNoProof(run);
		stylingDoc.FontFamilySize(run);
		Statsparagraph.setNumID(stylingDoc.bullet(document));
		logger.info("Writing All Campaigns count in doc file.....");
		run.setText("They have " + passData.Exceldata("All Campaigns") + " campaigns");

		Statsparagraph = document.createParagraph();
		Statsparagraph.setSpacingAfter(0);
		run = Statsparagraph.createRun();
		stylingDoc.FontFamilySize(run);
		Statsparagraph.setNumID(stylingDoc.bullet(document));
		logger.info("Writing Active Campaigns count in doc file.....");
		run.setText("They have " + passData.Exceldata("Active Campaigns") + " active campaigns");

		Statsparagraph = document.createParagraph();
		run = Statsparagraph.createRun();
		Statsparagraph.setSpacingAfter(0);
		stylingDoc.FontFamilySize(run);
		Statsparagraph.setNumID(stylingDoc.bullet(document));
		logger.info("Writing All Triggered Campaigns count in doc file.....");
		run.setText("They have " + passData.Exceldata("All Triggered Campaigns") + " triggered campaigns");

		Statsparagraph = document.createParagraph();
		run = Statsparagraph.createRun();
		Statsparagraph.setSpacingAfter(0);
		stylingDoc.FontFamilySize(run);
		Statsparagraph.setNumID(stylingDoc.bullet(document));
		logger.info("Writing Active Triggered Campaigns count in doc file.....");
		run.setText("They have " + passData.Exceldata("Active Triggered Campaigns") + " Active triggered campaigns");

		Statsparagraph = document.createParagraph();
		run = Statsparagraph.createRun();
		Statsparagraph.setSpacingAfter(0);
		stylingDoc.FontFamilySize(run);
		Statsparagraph.setNumID(stylingDoc.bullet(document));
		logger.info("Writing Batch Campaigns - Repeating Schedule count in doc file.....");
		run.setText("They have " + passData.Exceldata("Batch Campaigns - Repeating Schedule")
				+ " re-occurring batch campaigns");

		Statsparagraph = document.createParagraph();
		run = Statsparagraph.createRun();
		stylingDoc.FontFamilySize(run);
		Statsparagraph.setSpacingAfter(0);
		Statsparagraph.setNumID(stylingDoc.bullet(document));
		logger.info("Writing All Batch Campaigns count in doc file.....");
		run.setText("They have " + passData.Exceldata("All Batch Campaigns") + " batch campaigns");

		Statsparagraph = document.createParagraph();
		run = Statsparagraph.createRun();
		stylingDoc.FontFamilySize(run);
		Statsparagraph.setSpacingAfter(0);
		Statsparagraph.setNumID(stylingDoc.bullet(document));
		logger.info("Writing Landing Pages count in doc file.....");
		run.setText(passData.Exceldata("Landing Pages") + " landing pages");

		Statsparagraph = document.createParagraph();
		run = Statsparagraph.createRun();
		stylingDoc.FontFamilySize(run);
		Statsparagraph.setSpacingAfter(0);
		Statsparagraph.setNumID(stylingDoc.bullet(document));
		logger.info("Writing Forms count in doc file.....");
		run.setText(passData.Exceldata("Forms") + " forms");

		Statsparagraph = document.createParagraph();
		run = Statsparagraph.createRun();
		stylingDoc.FontFamilySize(run);
		Statsparagraph.setSpacingAfter(0);
		Statsparagraph.setNumID(stylingDoc.bullet(document));
		logger.info("Writing Emails count in doc file.....");
		run.setText(passData.Exceldata("Emails") + " emails");

		Statsparagraph = document.createParagraph();
		run = Statsparagraph.createRun();
		stylingDoc.FontFamilySize(run);
		Statsparagraph.setSpacingAfter(0);
		Statsparagraph.setNumID(stylingDoc.bullet(document));
		logger.info("Writing Snippets count in doc file.....");
		run.setText(passData.GenricMethod(passData.Exceldata("Snippets")));

		Statsparagraph = document.createParagraph();
		run = Statsparagraph.createRun();
		stylingDoc.FontFamilySize(run);
		Statsparagraph.setSpacingAfter(0);
		Statsparagraph.setNumID(stylingDoc.bullet(document));
		logger.info("Writing uploaded files count in doc file.....");
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

		try {
			int Snippets = Integer.parseInt(passData.Exceldata("Snippets"));
			if (Snippets < 5) {
				logger.info("Cont is less then 5 and it printing corresponding text");
				XWPFParagraph paragraphModel = document.createParagraph();
				XWPFRun model = paragraphModel.createRun();
				model.setBold(true);
				model.addCarriageReturn();
				model.setText("Snippets");
				
				XWPFParagraph SnippetsData = document.createParagraph();
				XWPFRun SnippetsRun = SnippetsData.createRun();
				stylingDoc.setNoProof(SnippetsRun);
				stylingDoc.FontFamilySize(SnippetsRun);
				SnippetsRun.setText(passData.Exceldata("Account Name") + passData.No_Snippets);

				XWPFParagraph paragraphModel2 = document.createParagraph();
				XWPFRun modellink = paragraphModel2.createRun();
				stylingDoc.setNoProof(modellink);
				modellink.setText(
						"\nhttps://experienceleague.adobe.com/docs/marketo/using/product-docs/personalization/segmentation-and-snippets/snippets/create-a-snippet.html?lang=en");
				paragraphModel.createHyperlinkRun(
						"https://experienceleague.adobe.com/docs/marketo/using/product-docs/personalization/segmentation-and-snippets/snippets/create-a-snippet.html?lang=en");
				modellink.setUnderline(UnderlinePatterns.SINGLE);
				modellink.setColor("3333cc");
				stylingDoc.FontFamilySize(modellink);
				logger.info("Stats Section completed");

			}}
			catch (Exception e) {
				logger.warn("Test is failed hence some data is missing from Stats Section\n");
				// TODO: handle exception
			}
		
		
		XWPFParagraph paragraphModel = document.createParagraph();
		XWPFRun model = paragraphModel.createRun();
		model.setBold(true);
		model.addCarriageReturn();
		model.setText("Programs");
		
		XWPFParagraph paragraphStats = document.createParagraph();
		XWPFRun r1 = paragraphStats.createRun();
		try {
			paragraphStats.setAlignment(ParagraphAlignment.LEFT);
			stylingDoc.setNoProof(r1);
			
			r1.setBold(false);

			stylingDoc.FontFamilySize(r1);
			r1.setText(passData.Exceldata("Account Name") + "," + passData.Org_info);

			XWPFRun r2 = paragraphStats.createRun();

			r2.setText("https://docs.marketo.com/display/public/DOCS/Smart+Campaigns");
			stylingDoc.setNoProof(r2);
			paragraphStats.createHyperlinkRun("https://docs.marketo.com/display/public/DOCS/Smart+Campaigns");
			r2.setUnderline(UnderlinePatterns.SINGLE);
			r2.setColor("3333cc");
			stylingDoc.FontFamilySize(r2);

			XWPFParagraph imgPara = document.createParagraph();
			XWPFRun img = imgPara.createRun();
			
			for(int i=1;i<2;i++) {
				
			img.addCarriageReturn();
			img.addPicture(new FileInputStream(passData.FetchScreenshot("Tags"+i)), Document.PICTURE_TYPE_PNG,
				passData.FetchScreenshot("Tags"+i), Units.toEMU(380), Units.toEMU(120));
			img.addCarriageReturn();
			}
		}
		
			
			catch (Exception e) {
				logger.warn("Test is failed hence some data is missing from Program Section\n");

			}

		}
			

	public static void models(XWPFDocument document) throws IOException, InvalidFormatException {
		// TODO Auto-generated method stub
		
		XWPFParagraph paragraphModel = document.createParagraph();
		XWPFRun model = paragraphModel.createRun();
		model.setBold(true);
		model.addBreak(BreakType.PAGE);
		model.addCarriageReturn();
		model.addCarriageReturn();
		logger.info("Printing Models Data");
		model.setText("Models");
		
		try {
			
		int Model =Integer.parseInt(passData.Exceldata("Models")) ;
		if (Model >0)
				{
			logger.info("Models data is greater then 0 hence printing appropriate data");
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
			logger.info("Printing images in doc");
			XWPFRun img1 = model_img.createRun();
			try {
			for(int i=1;i<3;i++) {	
			img1.addPicture(new FileInputStream(passData.FetchScreenshot("Models"+i)), Document.PICTURE_TYPE_PNG, passData.FetchScreenshot("Models"+i),
					Units.toEMU(230), Units.toEMU(150));
	
			}}
			catch (Exception e) {
				// TODO: handle exception
				logger.warn("Test is failed to retrive models images\n");
			}
			
			
			
		}
		
		 
		else
		{
			logger.info("you hadn't added any models to your subscription " +passData.Exceldata("Account Name"));
			XWPFParagraph paragraphModelData = document.createParagraph();
			XWPFRun modelTest = paragraphModelData.createRun();
			
			stylingDoc.setNoProof(modelTest);
			modelTest.setBold(true);
			modelTest.setFontSize(11);
			
			modelTest.setText(passData.Exceldata("Account Name") + passData.No_models);
			logger.info("Model section completed");
			
		}
		}
		catch (Exception e) {
			logger.error("Test is failed hence does not found any data it having null values\n");
			// TODO: handle exception
		}
		
	}
	
	public static void lead(XWPFDocument document) throws IOException, XmlException
	{
		XWPFParagraph LeadParagraph = document.createParagraph();
		XWPFRun LeadRun = LeadParagraph.createRun();
		
		LeadRun.addCarriageReturn();
		LeadRun.setBold(true);
		logger.info("Printing Lead scoring Section");
		LeadRun.setText("Lead Scoring");
		
		
		
		// adding lead scoring bullet point data into report
		XWPFParagraph LeadData = document.createParagraph();
		
		XWPFRun LeadDataRun = LeadData.createRun();
		stylingDoc.FontFamilySize(LeadDataRun);
		stylingDoc.setNoProof(LeadDataRun);
		LeadData.setSpacingAfter(0);
		LeadData.setNumID(stylingDoc.bullet(document));
		LeadDataRun.setText(passData.LeadScoring(passData.Exceldata("Change Score")));
		
		try {
		int ChangeScore = Integer.parseInt(passData.Exceldata("Change Score"));
		if (ChangeScore>1) 	
			logger.info("Change Score count is greater then 1");	
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
		}

			catch(NumberFormatException ex){
				
				logger.error("Lead Scoring having Null Data, So this section skipped\n");
				// handle your exception
		    
		}
		logger.info("Completed Lead Scoring Section");
	}

	public static void intrestingMoment(XWPFDocument document) throws NumberFormatException, InvalidFormatException, FileNotFoundException, IOException
	{
		
		if(passData.Exceldata("Interesting Moment Subscription").equalsIgnoreCase("true"))
		{
			logger.info("User have Interesting Moment Subscription");
			XWPFParagraph Interesting = document.createParagraph();
			XWPFRun InterestingRun = Interesting.createRun();
			InterestingRun.setBold(true);
			InterestingRun.addCarriageReturn();
			InterestingRun.addCarriageReturn();
			InterestingRun.setText("Interesting Moment");
			try {
			int Interesting_Moment= Integer.parseInt(passData.Exceldata("Interesting Moment"));
			
			if (Interesting_Moment>0)
			{
			
			XWPFParagraph InterestingData = document.createParagraph();
			XWPFRun InterestingDatarun = InterestingData.createRun();
			
			stylingDoc.setNoProof(InterestingDatarun);
			stylingDoc.FontFamilySize(InterestingDatarun);
			InterestingDatarun.setText("The following Interesting Moments have been defined to support\n"+passData.intresting_moment);
			
			InterestingDatarun = InterestingData.createRun();
			for(int i=1;i<4;i++) {
				try {
				InterestingDatarun.addCarriageReturn();
			InterestingDatarun.addPicture(new FileInputStream(passData.FetchScreenshot("Interesting Moment"+i)), Document.PICTURE_TYPE_PNG, passData.FetchScreenshot("Interesting Moment"),
					Units.toEMU(470), Units.toEMU(80));
			InterestingDatarun.addCarriageReturn();
			}
				catch (Exception e) {
					
					// TODO: handle exception
				}
			}
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
			
		catch(NumberFormatException ex){ // handle your exception
			logger.info("Test is skipped or fail hence data is missing for Intresting Moment\n");
		}
		}
		logger.info("Intresting Moment part is done");
	}

	public static void DataManagment(XWPFDocument document) throws IOException
	{
		XWPFParagraph DataManagement = document.createParagraph();
		XWPFRun DataManagementRun = DataManagement.createRun();
		DataManagementRun.setBold(true);
		DataManagementRun.addCarriageReturn();
		DataManagementRun.addCarriageReturn();
		logger.info("Printing data of Data Management...");
		DataManagementRun.setText("Data Management");
		
		
		XWPFParagraph DataManagementData = document.createParagraph();
		XWPFRun DataManagementDatarun = DataManagementData.createRun();
		stylingDoc.FontFamilySize(DataManagementDatarun);
		stylingDoc.setNoProof(DataManagementDatarun);
		DataManagementDatarun.setText(passData.DataManagement());	
	}
	
	public static void Events(XWPFDocument document) throws IOException
	{
		
		XWPFParagraph Events = document.createParagraph();
		XWPFRun	Eventsrun = Events.createRun();
		Eventsrun.setBold(true);
		Eventsrun.addCarriageReturn();
		logger.info("Event Data is printing");
		Eventsrun.setText("Events");

		XWPFParagraph EventsData = document.createParagraph();
		XWPFRun EventsDatarun = EventsData.createRun();
		stylingDoc.FontFamilySize(EventsDatarun);
		stylingDoc.setNoProof(EventsDatarun);
		EventsDatarun.setText(passData.Events());
		
		XWPFRun EventsDatarun2 = EventsData.createRun();
		stylingDoc.FontFamilySize(EventsDatarun2);
		stylingDoc.setNoProof(EventsDatarun2);
		EventsDatarun2.addCarriageReturn();
		EventsDatarun2.addCarriageReturn();
		EventsDatarun2.setText(passData.event2);
		
		XWPFRun EventsDatarunLink = EventsData.createRun();
		EventsDatarunLink.setUnderline(UnderlinePatterns.SINGLE);
		stylingDoc.FontFamilySize(EventsDatarunLink);
		stylingDoc.setNoProof(EventsDatarunLink);
		EventsDatarunLink.setUnderlineColor("3333cc");
		EventsDatarunLink.setColor("3333cc");
		EventsDatarunLink.setText("\nhttps://docs.marketo.com/display/public/DOCS/Create+a+New+Event+Program.");
		logger.info("Events data printing  is done");
	}
	
	public static void nurtureData(XWPFDocument document) throws NumberFormatException, IOException, XmlException

	
	{
		logger.info("Printing Nurture data");
		XWPFParagraph NurtureData = document.createParagraph();
		XWPFRun NurtureDataRun = NurtureData.createRun();
		NurtureDataRun.setBold(true);
		NurtureDataRun.addCarriageReturn();
		NurtureDataRun.setText("Nurture");
		;
		
		try {
		int NurtureDataint= Integer.parseInt(passData.Exceldata("Nurture campaigns"));
		
		if (NurtureDataint>0)
			
		{
		NurtureDataRun = NurtureData.createRun();
		stylingDoc.FontFamilySize(NurtureDataRun);
		stylingDoc.setNoProof(NurtureDataRun);
		NurtureDataRun.addCarriageReturn();
		NurtureDataRun.setText(passData.Exceldata("Account Name")+"\nhas\n"+passData.Exceldata("Nurture campaigns")+passData.Nurture);
		
		
		XWPFParagraph NurtureData1 = document.createParagraph();
		XWPFRun	NurtureDatarun1 = NurtureData1.createRun();
		stylingDoc.FontFamilySize(NurtureDatarun1);
		NurtureDatarun1.setBold(true);
		NurtureData1.setNumID(stylingDoc.bullet(document));
		
		NurtureDatarun1.setText("Intelligently and automatically deliver content to a target audience.");
		
		
		XWPFParagraph NurtureData2 = document.createParagraph();
		XWPFRun NurtureDatarun2 = NurtureData2.createRun();
		stylingDoc.FontFamilySize(NurtureDatarun2);
		NurtureDatarun2.setBold(true);
		
		NurtureData2.setNumID(stylingDoc.bullet(document));
		NurtureDatarun2.setText("Easily build dialogue with prospects and customers while preventing customers who have already received content from receiving the same content again.");
		
		XWPFParagraph NurtureData3 = document.createParagraph();
		XWPFRun	NurtureDatarun3 = NurtureData3.createRun();
		stylingDoc.FontFamilySize(NurtureDatarun3);
		NurtureDatarun3.setBold(true);
		
		NurtureData3.setNumID(stylingDoc.bullet(document));
		NurtureDatarun3.setText("Add new content and entire programs to nurture streams.");
		
		XWPFParagraph NurtureData4 = document.createParagraph();
		XWPFRun NurtureDatarun4 = NurtureData4.createRun();
		stylingDoc.FontFamilySize(NurtureDatarun4);
		NurtureDatarun4.setBold(true);
		
		NurtureData4.setNumID(stylingDoc.bullet(document));
		NurtureDatarun4.setText("Edit the availability of content.");
		
		XWPFParagraph NurtureData5 = document.createParagraph();
		XWPFRun	NurtureDatarun5 = NurtureData5.createRun();
		stylingDoc.FontFamilySize(NurtureDatarun5);
		NurtureDatarun5.setBold(true);
		NurtureData5.setNumID(stylingDoc.bullet(document));
		NurtureDatarun5.setText("Understand content performance based on engagement with each piece of content.");
		
		
		XWPFParagraph NurtureData6 = document.createParagraph();
		XWPFRun	NurtureDatarun6 = NurtureData6.createRun();
		stylingDoc.FontFamilySize(NurtureDatarun6);
		NurtureDatarun6.setText(passData.Nurture2);
		
		XWPFRun	NurtureDatarunLink = NurtureData6.createRun();
		stylingDoc.FontFamilySize(NurtureDatarunLink);
		stylingDoc.setNoProof(NurtureDatarunLink);
		NurtureDatarunLink.setUnderline(UnderlinePatterns.SINGLE);
		NurtureDatarunLink.setColor("3333cc");
		NurtureDatarunLink.setText("\nhttps://experienceleague.adobe.com/docs/marketo/using/product-docs/email-marketing/drip-nurturing/creating-an-engagement-program/create-an-engagement-program.html?lang=en .”");
		}
		
		
		else
		{
			NurtureDataRun = NurtureData.createRun();
			stylingDoc.FontFamilySize(NurtureDataRun);
			stylingDoc.setNoProof(NurtureDataRun);
			NurtureDataRun.addCarriageReturn();
			NurtureDataRun.setText(passData.Exceldata("Account Name")+passData.No_Nurture);
			
			XWPFRun NurtureDataRunlink = NurtureData.createRun();
			stylingDoc.FontFamilySize(NurtureDataRunlink);
		//	NurtureData.createHyperlinkRun("https://experienceleague.adobe.com/docs/marketo/using/product-docs/email-marketing/drip-nurturing/creating-an-engagement-program/create-an-engagement-program.html?lang=en\r\n");
			NurtureDataRunlink.setUnderline(UnderlinePatterns.SINGLE);
			NurtureDataRunlink.setColor("3333cc");
			NurtureDataRunlink.setText("https://experienceleague.adobe.com/docs/marketo/using/product-docs/email-marketing/drip-nurturing/creating-an-engagement-program/create-an-engagement-program.html?lang=en");
			
			
		
		
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
		
		}
		logger.info(" Nurture data part is done......");
		}catch(NumberFormatException ex){
			logger.error("Nurture Data Test is failed Hence we have null value\n");
			// handle your exception
		    
		}
		
	}

	public static void segment(XWPFDocument document) throws NumberFormatException, IOException, InvalidFormatException
	{
		XWPFParagraph Segment = document.createParagraph();
		XWPFRun SegmentHeading = Segment.createRun();
		SegmentHeading.setBold(true);
		logger.info("Printing Segmentation Data");
		SegmentHeading.setText("Segmentation");
		try {
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
	
		for(int i=1; i<4;i++) {
			try {
			SegmentRun.addCarriageReturn();
		SegmentRun.addPicture(new FileInputStream(passData.FetchScreenshot("Segmentations"+i)), Document.PICTURE_TYPE_PNG, passData.FetchScreenshot("Segmentations"),
				Units.toEMU(230), Units.toEMU(50));
		SegmentRun.addCarriageReturn();
		}
			catch (Exception e) {
				// TODO: handle exception
				logger.warn("We dont find any images so it does not printing images\n");
			}}
				}
		else {
			
			XWPFParagraph InterestingData = document.createParagraph();
			XWPFRun InterestingDatarun = InterestingData.createRun();
			stylingDoc.FontFamilySize(InterestingDatarun);
			stylingDoc.setNoProof(InterestingDatarun);
			InterestingDatarun.setText(passData.Exceldata("Account Name")+passData.no_segment);
			
			
			
		}
		}
		catch(NumberFormatException ex){ // handle your exception
		    logger.error("Test is skipped or failed so we have null value here\n");
		}
		logger.info(" Segmentation data part is done......");
	}
	
	public static void programLibrary(XWPFDocument document) throws IOException, InvalidFormatException
	{
		XWPFParagraph program = document.createParagraph();
		XWPFRun programRun = program.createRun();
		programRun.addCarriageReturn();
		programRun.addCarriageReturn();
		programRun.setBold(true);
		logger.info(" Program Library part is Started");
		programRun.setText("Program Library");
		
		
		
		XWPFParagraph ProgramData = document.createParagraph();
		XWPFRun ProgramDataRun = ProgramData.createRun();
		stylingDoc.FontFamilySize(ProgramDataRun);
		ProgramDataRun.addCarriageReturn();
		stylingDoc.setNoProof(ProgramDataRun);
		ProgramDataRun.setText(passData.Library());
		
	}

	public static void Integration(XWPFDocument document) throws IOException, InvalidFormatException
	{
		XWPFParagraph intigration = document.createParagraph();
		XWPFRun intigrationRunHeading = intigration.createRun();
		intigrationRunHeading.addCarriageReturn();
		intigrationRunHeading.addCarriageReturn();
		intigrationRunHeading.setBold(true);
		logger.info("Printing Integration Data");
		intigrationRunHeading.setText("Integrations");
		try {
		int intigration_Data= Integer.parseInt(passData.Exceldata("Integration"));
		if (intigration_Data>0)
		{
		
		XWPFParagraph intigrationRun = document.createParagraph();
		XWPFRun intigrationData = intigrationRun.createRun();
		stylingDoc.FontFamilySize(intigrationData);
		intigrationData.setText("The following integrations have been installed:");

		intigrationRun = document.createParagraph();
		XWPFRun intigrationImg = intigrationRun.createRun();
		for(int i =1; i<2;i++) {
			try {
			intigrationImg.addCarriageReturn();
		intigrationImg.addPicture(new FileInputStream(passData.FetchScreenshot("Integration")), Document.PICTURE_TYPE_PNG, passData.FetchScreenshot("Integration"),
				Units.toEMU(470), Units.toEMU(200));
		intigrationImg.addCarriageReturn();
		}
			catch (Exception e) {
				logger.warn("we dont have images to print\n");
				// TODO: handle exception
			}
		}
		
		}
		else {
			
			XWPFParagraph intigrationRun = document.createParagraph();
			XWPFRun intigrationData = intigrationRun.createRun();
			stylingDoc.FontFamilySize(intigrationData);
			stylingDoc.setNoProof(intigrationData);
			intigrationData.setText(passData.Exceldata("Account Name")+passData.No_Integrations);
			
			
		}
		}
		catch(NumberFormatException ex){ // handle your exception
		  logger.error("Intigration Test is skiped\n");  
		}
	
	}

	public static void webPersonalize(XWPFDocument document) throws IOException, InvalidFormatException
	{
		if(passData.Exceldata("Web Personalize").equalsIgnoreCase("true"))
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
		
		
		XWPFRun PersonalizeImg1 = PersonalizeDoc.createRun();
		stylingDoc.FontFamilySize(PersonalizeImg1);
		stylingDoc.setNoProof(PersonalizeImg1);
		PersonalizeImg1.addCarriageReturn();
		PersonalizeImg1.addPicture(new FileInputStream(passData.FetchScreenshot("Change Score1")), Document.PICTURE_TYPE_PNG, passData.FetchScreenshot("Change Score"),
				Units.toEMU(470), Units.toEMU(80));
		PersonalizeImg1.addBreak();
		PersonalizeImg1.addCarriageReturn();
		PersonalizeImg1.addPicture(new FileInputStream(passData.FetchScreenshot("Change Score1")), Document.PICTURE_TYPE_PNG, passData.FetchScreenshot("Change Score"),
				Units.toEMU(470), Units.toEMU(80));
		PersonalizeImg1.addBreak();
		PersonalizeImg1.addCarriageReturn();
		PersonalizeImg1.addPicture(new FileInputStream(passData.FetchScreenshot("Change Score1")), Document.PICTURE_TYPE_PNG, passData.FetchScreenshot("Change Score"),
				Units.toEMU(470), Units.toEMU(80));
		PersonalizeImg1.addCarriageReturn();
		PersonalizeImg1.setText("Total Organizations & Top 5 Organizations: The Organizations tab displays all the details (name, location, activity and time stamp) of organizations that visited your"
				+ " website during a given period. The table can be sorted and organized by time, location, domain and via a free text search.");
		
		//PersonalizeImg1.addCarriageReturn();
		PersonalizeImg1.addPicture(new FileInputStream(passData.FetchScreenshot("Change Score1")), Document.PICTURE_TYPE_PNG, passData.FetchScreenshot("Change Score"),
				Units.toEMU(470), Units.toEMU(80));
		PersonalizeImg1.addCarriageReturn();
		
		}
	}

	public static void TAM(XWPFDocument document) throws IOException, InvalidFormatException
	{
		if(passData.Exceldata("Target Account Management").equalsIgnoreCase("true"))
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
		OverviewRun.addBreak();
		try {
		OverviewRun.addPicture(new FileInputStream(passData.FetchScreenshot("Integration")), Document.PICTURE_TYPE_PNG, passData.FetchScreenshot("Integration").toString(),
				Units.toEMU(470), Units.toEMU(80));
		OverviewRun.addBreak();
		OverviewRun.setText("Top Named Accounts (by Pipeline)");
		OverviewRun.addBreak();
		OverviewRun.addPicture(new FileInputStream(passData.FetchScreenshot("Integration")), Document.PICTURE_TYPE_PNG, passData.FetchScreenshot("Integration").toString(),
				Units.toEMU(470), Units.toEMU(280));
		}
		catch (Exception e) {
			logger.warn("Test is skipped so user dont have images to print\n");
			// TODO: handle exception
		}
		XWPFParagraph HigestAccount = document.createParagraph();
		XWPFRun HigestAccountRunHead = HigestAccount.createRun();
	
		HigestAccountRunHead.setBold(true);
		HigestAccountRunHead.setText("Highest Account Score");
		
		XWPFRun HigestAccountRunData = HigestAccount.createRun();
		stylingDoc.FontFamilySize(HigestAccountRunData);
		stylingDoc.setNoProof(HigestAccountRunData);
		HigestAccountRunHead.addCarriageReturn();
		HigestAccountRunData.setText("Marketo’s ABM Account Score is a systematic approach designed to help Sales and Marketing teams identify and prioritize "
				+ "the companies (including prospects) that are most likely to make a purchase.");
		HigestAccountRunData.addCarriageReturn();
		HigestAccountRunData.addCarriageReturn();
		
		HigestAccountRunData.setText("In the complex world of B2B buying processes, it’s rare that a single individual makes a purchase decision. "
				+ "There are often various roles involved, each with their own needs. Account-based scoring takes this into account by aggregating the"
				+ " lead scores from multiple leads and providing a score at an account level.\n" +passData.Exceldata("Account Name") +"’s highest account score is from Cisco with a\n" + "989." );
		XWPFRun HigestAccountRunimg = HigestAccount.createRun();
		HigestAccountRunimg.addPicture(new FileInputStream(passData.FetchScreenshot("Integration")), Document.PICTURE_TYPE_PNG, passData.FetchScreenshot("Integration").toString(),
				Units.toEMU(470), Units.toEMU(110));
		HigestAccountRunimg.setBold(true);
		HigestAccountRunimg.setText("Account Lists");
		try {
		HigestAccountRunimg.addCarriageReturn();
		HigestAccountRunimg.addPicture(new FileInputStream(passData.FetchScreenshot("Integration")), Document.PICTURE_TYPE_PNG, passData.FetchScreenshot("Integration").toString(),
				Units.toEMU(470), Units.toEMU(110));
		}
		catch (Exception e) {
			logger.warn("Test is skipped so user dont have images to print\n");
			// TODO: handle exception
		}
		XWPFParagraph LargestAccountList = document.createParagraph();
		XWPFRun LargestAccountListRun = LargestAccountList.createRun();
		LargestAccountListRun.setBold(true);
		LargestAccountListRun.setText("Largest Account List");
		
		
		XWPFRun LargestAccountListRunData = LargestAccountList.createRun();
		LargestAccountListRunData.addCarriageReturn();
		stylingDoc.FontFamilySize(LargestAccountListRunData);
		stylingDoc.setNoProof(LargestAccountListRunData);
		LargestAccountListRunData.setText("An account list is a collection of named accounts that can be targeted together. Account lists allow you to target named accounts by industry, location or size of the company. Marketo allows for up to 2,000 named accounts to be added to an account list.\n" +passData.Exceldata("Account Name")+"\n’s largest account list is NA Target Accounts, containing 545 named accounts and is responsible for $23.1 million in pipeline.");
		
		XWPFRun LargestAccountListRunimg = LargestAccountList.createRun();
		try {
		LargestAccountListRunimg.addPicture(new FileInputStream(passData.FetchScreenshot("Integration")), Document.PICTURE_TYPE_PNG, passData.FetchScreenshot("Integration").toString(),
				Units.toEMU(470), Units.toEMU(80));
		}
		catch (Exception e) {
			logger.warn("Test is skipped so user dont have images to print\n");
			// TODO: handle exception
		}
		LargestAccountListRunimg.addCarriageReturn();
		LargestAccountListRunimg.addCarriageReturn();
		stylingDoc.FontFamilySize(LargestAccountListRunimg);
		stylingDoc.setNoProof(LargestAccountListRunimg);
		LargestAccountListRunimg.setText("To learn more about using Account Lists in Marketo, visit: http://docs.marketo.com/display/public/DOCS/Account+Lists");

		}
		
	}

	public static void PredictiveContent(XWPFDocument document) throws IOException, InvalidFormatException
	{
		if(passData.Exceldata("Predictive Content").equalsIgnoreCase("true"))
		{
		XWPFParagraph Target_Account = document.createParagraph();
		XWPFRun Target_AccountRun = Target_Account.createRun();
		
		Target_AccountRun.addCarriageReturn();
		Target_AccountRun.addCarriageReturn();
		Target_AccountRun.setBold(true);
		Target_AccountRun.setText("Predictive Content");
		
		
		
		XWPFParagraph Account_Management = document.createParagraph();
		XWPFRun Account_ManagementRun = Account_Management.createRun();
		stylingDoc.FontFamilySize(Account_ManagementRun);
		stylingDoc.setNoProof(Account_ManagementRun);
		Account_ManagementRun.setText(passData.PredictiveContent);
		
		
	
		}
		
	}

	
	
}


	
	



	
