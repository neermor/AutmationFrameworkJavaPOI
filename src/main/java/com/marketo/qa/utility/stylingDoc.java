package com.marketo.qa.utility;


import java.io.IOException;
import java.math.BigInteger;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFAbstractNum;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFNumbering;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.xmlbeans.XmlException;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTAbstractNum;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTNumbering;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTText;

public class stylingDoc {
	
	//bullets and Decimal values 
	
	static String cTAbstractNumBulletXML = 
	        "<w:abstractNum xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" w:abstractNumId=\"0\">"
	                + "<w:multiLevelType w:val=\"hybridMultilevel\"/>"
	                + "<w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"bullet\"/><w:lvlText w:val=\"\"/><w:lvlJc w:val=\"left\"/><w:pPr><w:ind w:left=\"720\" w:hanging=\"360\"/></w:pPr><w:rPr><w:rFonts w:ascii=\"Symbol\" w:hAnsi=\"Symbol\" w:hint=\"default\"/></w:rPr></w:lvl>"
	                + "<w:lvl w:ilvl=\"1\" w:tentative=\"1\"><w:start w:val=\"1\"/><w:numFmt w:val=\"bullet\"/><w:lvlText w:val=\"o\"/><w:lvlJc w:val=\"left\"/><w:pPr><w:ind w:left=\"1440\" w:hanging=\"360\"/></w:pPr><w:rPr><w:rFonts w:ascii=\"Courier New\" w:hAnsi=\"Courier New\" w:cs=\"Courier New\" w:hint=\"default\"/></w:rPr></w:lvl>"
	                + "<w:lvl w:ilvl=\"2\" w:tentative=\"1\"><w:start w:val=\"1\"/><w:numFmt w:val=\"bullet\"/><w:lvlText w:val=\"\"/><w:lvlJc w:val=\"left\"/><w:pPr><w:ind w:left=\"2160\" w:hanging=\"360\"/></w:pPr><w:rPr><w:rFonts w:ascii=\"Wingdings\" w:hAnsi=\"Wingdings\" w:hint=\"default\"/></w:rPr></w:lvl>"
	                + "</w:abstractNum>";

	static String cTAbstractNumDecimalXML = 
	        "<w:abstractNum xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" w:abstractNumId=\"0\">"
	                + "<w:multiLevelType w:val=\"hybridMultilevel\"/>"
	                + "<w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"decimal\"/><w:lvlText w:val=\"%1\"/><w:lvlJc w:val=\"left\"/><w:pPr><w:ind w:left=\"720\" w:hanging=\"360\"/></w:pPr></w:lvl>"
	                + "<w:lvl w:ilvl=\"1\" w:tentative=\"1\"><w:start w:val=\"1\"/><w:numFmt w:val=\"decimal\"/><w:lvlText w:val=\"%1.%2\"/><w:lvlJc w:val=\"left\"/><w:pPr><w:ind w:left=\"1440\" w:hanging=\"360\"/></w:pPr></w:lvl>"
	                + "<w:lvl w:ilvl=\"2\" w:tentative=\"1\"><w:start w:val=\"1\"/><w:numFmt w:val=\"decimal\"/><w:lvlText w:val=\"%1.%2.%3\"/><w:lvlJc w:val=\"left\"/><w:pPr><w:ind w:left=\"2160\" w:hanging=\"360\"/></w:pPr></w:lvl>"
	                + "</w:abstractNum>";

	//End of decimal values 
	
	
	public static BigInteger bullet(XWPFDocument document) throws IOException, XmlException
	{
		//XWPFDocument document;

		Path screenshotsDocumentPath = Paths.get(reports.screenshotsPathWordDoc);
//		if (!Files.exists(screenshotsDocumentPath)) {
//			// Create a blank document
//			document = new XWPFDocument();
//		} else {
//			// Open existing document
//			document = new XWPFDocument(Files.newInputStream(Paths.get(reports.screenshotsPathWordDoc)));
//		}

		//XWPFParagraph paragraph = document.createParagraph();

		//XWPFRun run = paragraph.createRun();
		//document = new XWPFDocument();
		CTNumbering cTNumbering = CTNumbering.Factory.parse(cTAbstractNumBulletXML);
        CTAbstractNum cTAbstractNum = cTNumbering.getAbstractNumArray(0);

        XWPFAbstractNum abstractNum = new XWPFAbstractNum(cTAbstractNum);
        XWPFNumbering numbering = document.createNumbering();
      BigInteger abstractNumID = numbering.addAbstractNum(abstractNum);

        BigInteger numID =numbering.addNum(abstractNumID);
       // System.out.println("numID: " + numID);
       // document.close();
        return numID;
        
	}
	
	
	 public static void HeaderFooter(XWPFDocument doc )
	    {
	    	
	    	CTSectPr sectPr=doc.getDocument().getBody().addNewSectPr();
	  		XWPFHeaderFooterPolicy policy=new XWPFHeaderFooterPolicy(doc,sectPr);
	  		
	  		// Creating Header objects
	  		CTP ctpHeader=CTP.Factory.newInstance();
	  		CTR ctrHeader=ctpHeader.addNewR();
	  		CTText ctHeader=ctrHeader.addNewT();
	  		
	  		// Alignments and color, styles of the header in the word doc -- Header
	  		XWPFParagraph headerParagraph=new XWPFParagraph(ctpHeader,doc);
	  		XWPFRun headerRun=headerParagraph.createRun();
	  		
	  		//Setting Header
	  		headerParagraph.setAlignment(ParagraphAlignment.CENTER);
	  		
	  		headerRun.addBreak();
	  		
	  		headerRun.setFontSize(15);
	  		headerRun.setColor("e60000");
	  		headerRun.setBold(true);
	  		headerRun.setText("Marketo Instance -Reports");
	  		headerRun.addTab();
	  		String curr_date=getCurrentDate("yyyy-MM-dd-hh:mm:ss");
	  		headerRun.setText(curr_date);
	  		
	  		//headerRun.setText(documentName);
	  		
	  		
	  		//Parse
	  		XWPFParagraph[] parsHeader=new XWPFParagraph[1];
	  		parsHeader[0]= headerParagraph;
	  		policy.createHeader(XWPFHeaderFooterPolicy.DEFAULT,parsHeader);
	  		
	  		//Create Footer Objects
	  		CTP ctpFooter=CTP.Factory.newInstance();
	  		CTR ctrFooter=ctpFooter.addNewR();
	  		CTText ctFooter=ctrFooter.addNewT();
	  		
	  		//Setting Footer
	  		Date date=new Date();
	  		int year=date.getYear();
	  		int currentYear=year+1900;  
	  		String footerText="©"+currentYear+". Confidential - Do not Share this documents.";
	  		ctFooter.setStringValue(footerText);	
	  		
	  		// Alignments and color, styles of the header in the word doc -- Footer
	  		XWPFParagraph footerparagraph = new XWPFParagraph(ctpFooter,doc);
	  		footerparagraph.setAlignment(ParagraphAlignment.LEFT);
	  		
	  		//Parse
	  		XWPFParagraph[] parsFooter=new XWPFParagraph[1];
	  		parsFooter[0]= footerparagraph;
	  		policy.createFooter(XWPFHeaderFooterPolicy.DEFAULT,parsFooter);//Corrected to Footer
	    
	    	
	    }
	 
	 private static String getCurrentDate(String format) 
		{
			DateFormat dateformat=new SimpleDateFormat(format);
			Date date=new Date();
			return dateformat.format(date);
			
		}


}
