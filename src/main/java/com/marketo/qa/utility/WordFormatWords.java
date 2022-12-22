package com.marketo.qa.utility;


import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import java.util.*;



	
	public class WordFormatWords {

	 static void cloneRunProperties(XWPFRun source, XWPFRun dest) { // clones the underlying w:rPr element
	  CTR tRSource = source.getCTR();
	  CTRPr rPrSource = tRSource.getRPr();
	  if (rPrSource != null) {
	   CTRPr rPrDest = (CTRPr)rPrSource.copy();
	   CTR tRDest = dest.getCTR();
	   tRDest.setRPr(rPrDest);
	  }
	 }

	 
	 static void formatWord(XWPFParagraph paragraph, String keyword, Map<String, String> formats) {
		 
	  int runNumber = 0;
	  while (runNumber < paragraph.getRuns().size()) { //go through all runs, we cannot use for each since we will possibly insert new runs
	   XWPFRun run = paragraph.getRuns().get(runNumber);
	   XWPFRun run2 = run;
	   String runText = run.getText(0);
	   if (runText != null && runText.contains(keyword)) { //if we have a run with keyword in it, then

	    char[] runChars = runText.toCharArray(); //split run text into characters
	    StringBuffer sb = new StringBuffer();
	    for (int charNumber = 0; charNumber < runChars.length; charNumber++) { //go through all characters in that run
	     sb.append(runChars[charNumber]); //buffer all characters
	     runText = sb.toString();
	     if (runText.endsWith(keyword)) { //if the bufferend character stream ends with the keyword  
	      //set all chars, which are current buffered, except the keyword, as the text of the actual run
	      run.setText(runText.substring(0, runText.length() - keyword.length()), 0); 
	      run2 = paragraph.insertNewRun(++runNumber); //insert new run for the formatted keyword
	      cloneRunProperties(run, run2); // clone the run properties from original run
	      run2.setText(keyword, 0); // set the keyword in run
	      for (String toSet : formats.keySet()) { // do the additional formatting
	       if ("color".equals(toSet)) {
	        run2.setColor(formats.get(toSet));
	       } else if ("bold".equals(toSet)) {
	        run2.setBold(Boolean.valueOf(formats.get(toSet)));
	       }
	      }
	      run2 = paragraph.insertNewRun(++runNumber); //insert a new run for the next characters
	      cloneRunProperties(run, run2); // clone the run properties from original run
	      run = run2;
	      sb = new StringBuffer(); //empty the buffer
	     } 
	    }
	    run.setText(sb.toString(), 0); //set all characters, which are currently buffered, as the text of the actual run

	   }
	   runNumber++;
	  }
	 }




}
