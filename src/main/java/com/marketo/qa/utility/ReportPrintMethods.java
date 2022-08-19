package com.marketo.qa.utility;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.BreakType;
import org.apache.poi.xwpf.usermodel.Document;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.xmlbeans.XmlException;

public class ReportPrintMethods {
	
	public static void stats(XWPFDocument document) throws IOException, XmlException 
	{
		XWPFParagraph Statsparagraph = document.createParagraph();
		XWPFRun run = Statsparagraph.createRun();
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
		run.setText(passData.GenricMethod(passData.Exceldata("Snippets")));
		System.out.println(passData.GenricMethod(passData.Exceldata("Snippets")));
		
		Statsparagraph = document.createParagraph();
		run = Statsparagraph.createRun();
		run.setFontFamily("Calibri Light (Headings)");
		run.setFontSize(10);
		Statsparagraph.setNumID(stylingDoc.bullet(document));
		run.setText(passData.Exceldata("Emails") + " emails");
		
		
		
		XWPFParagraph paragraphStats = document.createParagraph();
		XWPFRun r1 = paragraphStats.createRun();
		
			paragraphStats.setAlignment(ParagraphAlignment.LEFT);
			stylingDoc.setNoProof(r1);
			r1.addCarriageReturn();
			r1.addCarriageReturn();
			r1.setBold(false);
			r1.addCarriageReturn();
			
			r1.setFontFamily("Calibri Light (Headings)");
			r1.setFontSize(10);
			r1.setText(passData.Exceldata("Account Name") +"," +passData.Org_info);
			
			XWPFParagraph imgPara = document.createParagraph();
			XWPFRun img = imgPara.createRun();
			try {
				img.addPicture(new FileInputStream(passData.FetchScreenshot("Tags")), Document.PICTURE_TYPE_PNG, passData.FetchScreenshot("Tags"),
						Units.toEMU(380), Units.toEMU(150));
			} catch (InvalidFormatException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (FileNotFoundException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			
			img.addBreak(BreakType.PAGE);
			img.addCarriageReturn();
			img.addCarriageReturn();
			
		
	}
	

}
