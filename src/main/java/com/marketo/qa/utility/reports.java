package com.marketo.qa.utility;

import static org.testng.Assert.assertTrue;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.HashMap;
import java.util.Map;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.BreakType;
import org.apache.poi.xwpf.usermodel.Document;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.xmlbeans.XmlException;

public class reports {

	passData pd = new passData();

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
		paragraph.setNumID(stylingDoc.bullet(document));
		run.setText("They have " + passData.Exceldata("Change Score")+ " triggered campaigns");
		System.out.println(passData.Exceldata("Interesting Moment"));
		paragraph = document.createParagraph();
		run = paragraph.createRun();
		paragraph.setNumID(stylingDoc.bullet(document));
		run.setText("They have " + passData.Document_name + " triggered campaigns");
		paragraph = document.createParagraph();
		paragraph.setAlignment(ParagraphAlignment.LEFT);
		XWPFRun r1 = paragraph.createRun();

		try {
			int width = 400;
			int height = 240;

			r1.addCarriageReturn();
			r1.addCarriageReturn();
			r1.setBold(false);
			r1.addCarriageReturn();
			r1.setFontFamily("Calibri Light (Headings)");
			r1.setFontSize(10);
			r1.setText(passData.Org_info);

			r1.addPicture(new FileInputStream(passData.login), Document.PICTURE_TYPE_PNG, passData.login,
					Units.toEMU(width), Units.toEMU(height));
			r1.addBreak(BreakType.PAGE);
			r1.addCarriageReturn();
			r1.addCarriageReturn();

			r1.addPicture(new FileInputStream(passData.after), Document.PICTURE_TYPE_PNG, passData.after,
					Units.toEMU(width), Units.toEMU(height));

		} catch (Exception e) {

		}
		stylingDoc.HeaderFooter(document);// Styling of Header and footer

		String[] keywords = new String[] { passData.Document_name, passData.test,passData.Exceldata("Interesting Moment") };
		Map<String, String> formats = new HashMap<String, String>();
		formats.put("bold", "true");
		formats.put("color", "000000");

		for (XWPFParagraph AllParagraph : document.getParagraphs()) { // go through all paragraphs
			for (String keyword : keywords) {

				WordFormatWords.formatWord(AllParagraph, keyword, formats);
			}
		}
		FileOutputStream out = new FileOutputStream(screenshotsPathWordDoc);

		document.write(out);
		out.close();

	}

}
