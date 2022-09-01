package com.marketo.qa.utility;

import static org.testng.Assert.assertTrue;
import java.io.File;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

public class reports {

	static String word = System.getProperty("user.home") + "\\Desktop\\Reports\\";

	public static void docs() throws Exception {
		Path outputDirectory = Paths.get(word);
		if (!Files.exists(outputDirectory)) {
			assertTrue(new File(String.valueOf(outputDirectory)).mkdirs(), "Unable to create output directory");

		}
		
		
		//word format data pasting 
		XWPFDocument document = new XWPFDocument();
			docReports.stats(document);

			// Models report part, Note : we have to added conditional based formatting
			docReports.models(document);

			// adding lead scoring data into report
			docReports.lead(document);

			// Interesting moment report part
			docReports.intrestingMoment(document);

			// adding Data management

			docReports.DataManagment(document);

			// adding Events data
			docReports.Events(document);

			// adding Nurture data
			docReports.nurtureData(document);

			// Segment Data printing
			docReports.segment(document);

			// program library section
			docReports.programLibrary(document);

			// Integration Data Section
			docReports.Integration(document);

			// Web Personalize Data Section
			docReports.webPersonalize(document);

			// Target Account Management
			docReports.TAM(document);

			docReports.PredictiveContent(document);
			//closing the document
			docReports.close(document);
	}

}
