package com.marketo.qa.utility;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.BreakType;
import org.apache.poi.xwpf.usermodel.Document;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTText;

public class docReports {

	static String word = System.getProperty("user.home") + "\\Desktop\\Reports\\";
	static String screenshotsPathWordDoc = word + "screenshots.docx";

	// Create Separate folder

	public static void CreateFolder() {
		File file1 = new File(word);
		if (!file1.exists()) {
			if (file1.mkdir()) {
				System.out.println("Directory is created!");
			} else {
				System.out.println("Failed to create directory!");
			}
		}
	}

	// defined header and footer

	public static void HeaderFooter(XWPFDocument doc) {

		CTSectPr sectPr = doc.getDocument().getBody().addNewSectPr();
		XWPFHeaderFooterPolicy policy = new XWPFHeaderFooterPolicy(doc, sectPr);

		// Creating Header objects
		CTP ctpHeader = CTP.Factory.newInstance();
		CTR ctrHeader = ctpHeader.addNewR();
		CTText tHeader = ctrHeader.addNewT();

		// Alignments and color, styles of the header in the word doc -- Header
		XWPFParagraph headerParagraph = new XWPFParagraph(ctpHeader, doc);
		XWPFRun headerRun = headerParagraph.createRun();

		// Setting Header
		headerParagraph.setAlignment(ParagraphAlignment.LEFT);

		headerRun.addBreak();

		headerRun.setFontSize(15);
		headerRun.setColor("e60000");
		headerRun.setBold(true);
		headerRun.setText("Marketo Instance -Reports");
		headerRun.addBreak();
		String curr_date = getCurrentDate("yyyy-MM-dd-hh:mm:ss");
		headerRun.setText(curr_date);
		headerRun.addBreak();
		// headerRun.setText(documentName);

		// Parse
		XWPFParagraph[] parsHeader = new XWPFParagraph[1];
		parsHeader[0] = headerParagraph;
		policy.createHeader(XWPFHeaderFooterPolicy.DEFAULT, parsHeader);

		// Create Footer Objects
		CTP ctpFooter = CTP.Factory.newInstance();
		CTR ctrFooter = ctpFooter.addNewR();
		CTText ctFooter = ctrFooter.addNewT();

		// Setting Footer
		String footerText = "© 2022. Confidential - Do not Share this documents.";
		ctFooter.setStringValue(footerText);

		// Alignments and color, styles of the header in the word doc -- Footer
		XWPFParagraph footerparagraph = new XWPFParagraph(ctpFooter, doc);
		footerparagraph.setAlignment(ParagraphAlignment.LEFT);

		// Parse
		XWPFParagraph[] parsFooter = new XWPFParagraph[1];
		parsFooter[0] = footerparagraph;
		policy.createFooter(XWPFHeaderFooterPolicy.DEFAULT, parsFooter);// Corrected to Footer

	}

	public static void data(String documentName, String[] screenshotNames) throws IOException, InvalidFormatException {
		CreateFolder();
		try (XWPFDocument doc = new XWPFDocument()) {
			XWPFParagraph p = doc.createParagraph();
			XWPFRun r = p.createRun();

			HeaderFooter(doc); // calling footer method
			p.setAlignment(ParagraphAlignment.RIGHT);

			r.setBold(true);
			r.setFontFamily("Verdana");

			r.addCarriageReturn();
			r.setText("Hello All demo :" + documentName + "Basic document");
			r.addBreak();
			r.addBreak();
			r.addBreak();
			for (String file : screenshotNames) {

				try {

					File dest = new File(
							System.getProperty("user.dir") + "\\test-output\\screenshots\\" + file + ".png");

					System.out.println(dest);
					FileInputStream is = new FileInputStream(dest);
					int width = 500;
					int height = 280;

					String imgFile = dest.getName();
					r.setText(file);
					r.addPicture(is, Document.PICTURE_TYPE_PNG, imgFile, Units.toEMU(width), Units.toEMU(height)); /* Writing
																													 images
																													 into
																													 word
																													 file
																													 with
																													 width
																													 and
																													 height*/
					r.addBreak(BreakType.PAGE);

				}

				catch (Exception e) {
					continue;
				}
			}
			saveDocs(doc);
		}

	}

//	public static void reports(String documentName, String[] screenshotNames) throws IOException, InvalidFormatException  {
//
//		CreateFolder();
//		
//		
//   
//    try (XWPFDocument doc = new XWPFDocument()) {
//    	System.out.println(doc);
//
//        XWPFParagraph p = doc.createParagraph();
//       
//        XWPFRun r = p.createRun();
//        HeaderFooter(doc);				//calling footer method 
//        
//        p.setAlignment(ParagraphAlignment.LEFT);
//		r.setBold(true);
//		r.setFontFamily("Verdana");
//		r.setText(documentName);
//		
//		
//		r.addBreak();
//		
//		BufferedImage bfImg;

//for (String file : screenshotNames )  {
//	
//		
//	
//			try {
//				
//					
//					File dest = new File(System.getProperty("user.dir")+"\\test-output\\screenshots\\" + file+ ".png");
//					//FileInputStream is = new FileInputStream(dest);
//					bfImg = ImageIO.read(dest);
//					FileInputStream is = new FileInputStream(dest);
//					int width = 500;
//					int height = 280;
//	
//					String imgFile = dest.getName();
//					
//					//int imgFormat = getImageFormat(imgFile);
//					
//					r.addBreak();
//					r.addBreak();
//					r.setText(file);				
//					r.addPicture(is, Document.PICTURE_TYPE_PNG,
//							imgFile, Units.toEMU(width), Units.toEMU(height)); //Writing images into word file with width and height 
//					     
//				}
//			catch(Exception e)
//			{
//				continue;
//			}
//			
//		
//			
//		}
//}
//		System.out.println("Word document with screenshots created successfully");
//		
//		saveDocs(doc);
//		doc.close();
//    }

	// Created saveDocs method Closing word document

	public static void saveDocs(XWPFDocument doc) throws IOException {
		FileOutputStream out = new FileOutputStream(screenshotsPathWordDoc);
		doc.write(out);
		out.close();

	}

	private static String getCurrentDate(String format) {
		DateFormat dateformat = new SimpleDateFormat(format);
		Date date = new Date();
		return dateformat.format(date);

	}

}
