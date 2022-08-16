package com.marketo.qa.utility;


import java.io.IOException;
import java.math.BigInteger;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;
import java.util.TimeZone;

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
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTOnOff;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTText;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STOnOff;

import com.marketo.qa.Pages.AdminPage;

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
	
	//static String AccountName = new AdminPage().AccountName();


	//End of decimal values 
	
	
	public static BigInteger bullet(XWPFDocument document) throws IOException, XmlException
	{
		
		CTNumbering cTNumbering = CTNumbering.Factory.parse(cTAbstractNumBulletXML);
        CTAbstractNum cTAbstractNum = cTNumbering.getAbstractNumArray(0);

        XWPFAbstractNum abstractNum = new XWPFAbstractNum(cTAbstractNum);
        XWPFNumbering numbering = document.createNumbering();
      BigInteger abstractNumID = numbering.addAbstractNum(abstractNum);

        BigInteger numID =numbering.addNum(abstractNumID);
        return numID;
        
	}
	
	
	 public static void HeaderFooter(XWPFDocument doc ) throws IOException
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
	  		setNoProof(headerRun);
	  		
	  		//Setting Header
	  		headerParagraph.setAlignment(ParagraphAlignment.CENTER);
	  		
	  		headerRun.addBreak();
	  		
	  		headerRun.setFontSize(15);
	  		headerRun.setColor("e60000");
	  		headerRun.setBold(true);
	  		headerRun.setText(passData.AccountName +"– Instance Review –");
	  		String curr_date=getCurrentDate(" dd-MM-yyyy");
	  		headerRun.setText(curr_date + "–"+getCurrentDate(" hh:mm ")+ getCurrentDate("a")+" MST" );
	  		
	  		
	  		
	  		
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
	  		@SuppressWarnings("deprecation")
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
	 
	 // Removing Spell checker 
	 public static void setNoProof (XWPFRun run) {
		    CTR ctR = run.getCTR();
		    CTRPr ctRPr = ctR.isSetRPr() ? ctR.getRPr() : ctR.addNewRPr();
		    if (!ctRPr.isSetNoProof()) {
		        // If the noProof property is missing, add it
		        ctRPr.addNewNoProof();
		    } else {
		        // If the noProof property is present, make sure it is not
		        // FALSE, OFF, or X_0
		        CTOnOff noProof = ctRPr.getNoProof();
		        if (noProof.isSetVal() &&
		                (noProof.getVal() == STOnOff.FALSE ||
		                noProof.getVal() == STOnOff.OFF ||
		                noProof.getVal() == STOnOff.X_0)) {
		            noProof.setVal(STOnOff.TRUE);
		        }
		    }
		}
	 
	 // Using bellow method bold the particular data 
	 
	 public static void bold(XWPFDocument document) throws IOException
		{
			String[] keywords = new String[] {String.valueOf(passData.Exceldata("All Campaigns")),String.valueOf(passData.Exceldata("Active campaigns")), 
					String.valueOf(passData.Exceldata("Landing Pages")),String.valueOf(passData.Exceldata("Active Triggered Campaigns")),
					String.valueOf(passData.Exceldata("Forms")),String.valueOf(passData.Exceldata("Triggered campaigns")),String.valueOf(passData.Exceldata("Landing Pages")),
					"Not","not",String.valueOf(passData.Exceldata("Snippets")),String.valueOf(passData.Exceldata("Emails")),
					String.valueOf(passData.Exceldata("Active Campaigns")),String.valueOf(passData.Exceldata("model")),String.valueOf(passData.Exceldata("Lead")),
					String.valueOf(passData.Exceldata("Change Data Value")),String.valueOf(passData.Exceldata("Event Programs")),String.valueOf(passData.Exceldata("Nurture campaigns")),
					String.valueOf(passData.Exceldata("Segment Data")),String.valueOf(passData.Exceldata("Library")),String.valueOf(passData.Exceldata("Integrations")),
					String.valueOf(passData.AccountName),String.valueOf(passData.Exceldata("All Batch Campaigns")),
					String.valueOf(passData.Exceldata("Models")),String.valueOf(passData.Exceldata("Leads"))};
		
			Map<String, String> formats = new HashMap<String, String>();
			formats.put("bold", "true");
			formats.put("color", "000000");
			

			for (XWPFParagraph AllParagraph : document.getParagraphs()) { // go through all paragraphs
				for (String keyword : keywords) {

					WordFormatWords.formatWord(AllParagraph, keyword, formats);
				}}
			}
	 
	 private static String getCurrentDate(String format) 
		{
			
			Date date=new Date();
			DateFormat dateformat=new SimpleDateFormat(format);
			dateformat.setTimeZone(TimeZone.getTimeZone("MST"));
			return dateformat.format(date);
			
		}
	 
	


}
