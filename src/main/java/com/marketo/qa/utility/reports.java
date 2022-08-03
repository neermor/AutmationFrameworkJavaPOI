package com.marketo.qa.utility;

import static org.testng.Assert.assertTrue;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.BreakType;
import org.apache.poi.xwpf.usermodel.Document;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.xmlbeans.XmlException;

import com.marketo.qa.Pages.AdminPage;

public class reports {

	passData pd = new passData();
	static String AccountName = stylingDoc.AccountName;
	

	static String word = System.getProperty("user.home") + "\\Desktop\\Reports\\";
	static String screenshotsPathWordDoc = word + "screenshots.docx";

	public static void docs() throws IOException, XmlException {
		Path outputDirectory = Paths.get(word);
		if (!Files.exists(outputDirectory)) {
			assertTrue(new File(String.valueOf(outputDirectory)).mkdirs(), "Unable to create output directory");

		}

		XWPFDocument document;
		document = new XWPFDocument();
		
		XWPFParagraph paragraph = document.createParagraph();
		XWPFRun run = paragraph.createRun();
		run.setBold(true);
		run.setText("Stats");

		paragraph = document.createParagraph();
		run = paragraph.createRun();
		run.setFontFamily("Calibri Light (Headings)");
		run.setFontSize(10);
		paragraph.setNumID(stylingDoc.bullet(document));
		run.setText("They have " + passData.Exceldata("All Campaigns")+ " campaigns");
		System.out.println(passData.Exceldata("Interesting Moment"));
		
		paragraph = document.createParagraph();
		run = paragraph.createRun();
		run.setFontFamily("Calibri Light (Headings)");
		run.setFontSize(10);
		paragraph.setNumID(stylingDoc.bullet(document));
		run.setText("They have " + passData.Exceldata("Active Campaigns") + " active campaigns");
		
		
		paragraph = document.createParagraph();
		run = paragraph.createRun();
		run.setFontFamily("Calibri Light (Headings)");
		run.setFontSize(10);
		paragraph.setNumID(stylingDoc.bullet(document));
		run.setText("They have " + passData.Exceldata("Active Triggered Campaigns") + " triggered campaigns");
		
		
		paragraph = document.createParagraph();
		run = paragraph.createRun();
		run.setFontFamily("Calibri Light (Headings)");
		run.setFontSize(10);
		paragraph.setNumID(stylingDoc.bullet(document));
		run.setText( passData.Exceldata("Landing Pages") + " landing pages");
		
		paragraph = document.createParagraph();
		run = paragraph.createRun();
		run.setFontFamily("Calibri Light (Headings)");
		run.setFontSize(10);
		paragraph.setNumID(stylingDoc.bullet(document));
		run.setText(passData.Exceldata("Forms")+" forms");
		
		paragraph = document.createParagraph();
		run = paragraph.createRun();
		run.setFontFamily("Calibri Light (Headings)");
		run.setFontSize(10);
		paragraph.setNumID(stylingDoc.bullet(document));
		run.setText(passData.Exceldata("Emails") + " emails");
		
		
		
		paragraph = document.createParagraph();
		XWPFRun r1 = paragraph.createRun();
		try {
			paragraph.setAlignment(ParagraphAlignment.LEFT);
			r1.addCarriageReturn();
			r1.addCarriageReturn();
			r1.setBold(false);
			r1.addCarriageReturn();
			r1.setFontFamily("Calibri Light (Headings)");
			r1.setFontSize(10);
			r1.setText(AccountName+"," +passData.Org_info);
			
			paragraph = document.createParagraph();
			run = paragraph.createRun();

			r1.addPicture(new FileInputStream(passData.Tag), Document.PICTURE_TYPE_PNG, passData.Tag,
					Units.toEMU(380), Units.toEMU(150));
			r1.addBreak(BreakType.PAGE);
			r1.addCarriageReturn();
			r1.addCarriageReturn();
			
			//Models report part , Note : we have to add conditional based formatting 
			run = paragraph.createRun();
			run.setBold(true);
			run.setText("Models");
			run.addBreak();
			run = paragraph.createRun();
			run.setFontFamily("Calibri Light (Headings)");
			run.setFontSize(10);
			run.setText(AccountName +"," +passData.models);
			run.addBreak(BreakType.PAGE);
			
			//adding lead scoring data into report 
			run = paragraph.createRun();
			run.setBold(true);
			run.setText("Models");
			
			
			
			// adding lead scoring bullet point data into report 
			paragraph = document.createParagraph();
			run = paragraph.createRun();
			run.setFontFamily("Calibri Light (Headings)");
			run.setFontSize(10);
			paragraph.setNumID(stylingDoc.bullet(document));
			run.setText(AccountName+","+" has advanced lead scoring built out taking into account behavior, demographics, successes and decay.");
			
			paragraph = document.createParagraph();
			run = paragraph.createRun();
			run.setFontFamily("Calibri Light (Headings)");
			run.setFontSize(10);
			paragraph.setNumID(stylingDoc.bullet(document));
			run.setText(AccountName+","+"is executing multiple score changes with single campaigns.");
			
			paragraph = document.createParagraph();
			run = paragraph.createRun();
			run.setFontFamily("Calibri Light (Headings)");
			run.setFontSize(10);
			paragraph.setNumID(stylingDoc.bullet(document));
			run.setText(AccountName+","+"is using MyTokens in their lead scoring campaigns which allows for a Marketer to quickly, and easily, control from a high level their lead change scores.");
			
			paragraph = document.createParagraph();
			run = paragraph.createRun();
			run.setFontFamily("Calibri Light (Headings)");
			run.setFontSize(10);
			paragraph.setNumID(stylingDoc.bullet(document));
			run.setText(AccountName+","+"has built out campaigns reducing lead scores when leads exhibit undesirable behavior.");
		
			
			
			paragraph = document.createParagraph();
			run = paragraph.createRun();
			run.setFontFamily("Calibri Light (Headings)");
			run.setFontSize(10);
			run.setText(passData.lead_scoring);
			run.addCarriageReturn();
			run.addCarriageReturn();
			run.addCarriageReturn();
			
			//Interesting moment report part 
			paragraph = document.createParagraph();
			run = paragraph.createRun();
			run.setBold(true);
			run.setText("Interesting Moments");
			run.addBreak();
			
			run = paragraph.createRun();
			run.setFontFamily("Calibri Light (Headings)");
			run.setFontSize(10);
			run.setText(passData.intresting_moment +" " + AccountName +"’s marketing campaigns."
					+" When a lead exhibits any of the below behavior, it will be documented and tracked.");
			
			run.addPicture(new FileInputStream(passData.intesting_moment_img), Document.PICTURE_TYPE_PNG, passData.Tag,
					Units.toEMU(470), Units.toEMU(80));
			run.setText(passData.intresting_moment_below);
			run.addCarriageReturn();
		
			
		//adding Data management 	
			
			paragraph = document.createParagraph();
			run = paragraph.createRun();
			run.setBold(true);
			run.setText("Data Management");
			run.addBreak();
			
			run = paragraph.createRun();
			run.setFontFamily("Calibri Light (Headings)");
			run.setFontSize(10);
			run.setText(AccountName + passData.data_management);	
			run.addCarriageReturn();
			
		//adding Events data 
			
			paragraph = document.createParagraph();
			run = paragraph.createRun();
			run.setBold(true);
			run.setText("Events");
			run.addBreak();
			
			run = paragraph.createRun();
			run.setFontFamily("Calibri Light (Headings)");
			run.setFontSize(10);
			run.setText(AccountName + " has built numerous Event campaigns in Marketo.");	
			run.addBreak();
			run.setText(passData.events);
		
		//adding Nurture data 
			paragraph = document.createParagraph();
			run = paragraph.createRun();
			run.setBold(true);
			run.setText("Nurture");
			run.addBreak();
			run = paragraph.createRun();
			run.setFontFamily("Calibri Light (Headings)");
			run.setFontSize(10);
			run.setText(AccountName + passData.Nurture);
			
		
			paragraph = document.createParagraph();
			run = paragraph.createRun();
			run.setFontFamily("Calibri Light (Headings)");
			run.setFontSize(10);
			paragraph.setNumID(stylingDoc.bullet(document));
			run.setText("Intelligently and automatically deliver content to a target audience.");
			
			paragraph = document.createParagraph();
			run = paragraph.createRun();
			run.setFontFamily("Calibri Light (Headings)");
			run.setFontSize(10);
			paragraph.setNumID(stylingDoc.bullet(document));
			run.setText("Easily build dialogue with prospects and customers while preventing customers who have already received content from receiving the same content again.");
			
			paragraph = document.createParagraph();
			run = paragraph.createRun();
			run.setFontFamily("Calibri Light (Headings)");
			run.setFontSize(10);
			paragraph.setNumID(stylingDoc.bullet(document));
			run.setText("Add new content and entire programs to nurture streams.");
			
			paragraph = document.createParagraph();
			run = paragraph.createRun();
			run.setFontFamily("Calibri Light (Headings)");
			run.setFontSize(10);
			paragraph.setNumID(stylingDoc.bullet(document));
			run.setText("Edit the availability of content.");
			
			paragraph = document.createParagraph();
			run = paragraph.createRun();
			run.setFontFamily("Calibri Light (Headings)");
			run.setFontSize(10);
			paragraph.setNumID(stylingDoc.bullet(document));
			run.setText("Understand content performance based on engagement with each piece of content.");
			
			
		} catch (Exception e) {

		}
		stylingDoc.HeaderFooter(document);// Styling of Header and footer
		
		stylingDoc.bold(document);
		FileOutputStream out = new FileOutputStream(screenshotsPathWordDoc);

		document.write(out);
		out.close();

	}
	
	
		
	}


