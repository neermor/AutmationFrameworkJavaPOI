package com.marketo.qa.utility;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.Document;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.UnderlinePatterns;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.xmlbeans.XmlException;

public class docReports {

	static String word = System.getProperty("user.dir") + "//Reports//";
	static String fileName = new SimpleDateFormat("dd_MM_yy_HH_mm").format(new Date());
	private static final Logger logger = LogManager.getLogger(docReports.class);

	public static void close(XWPFDocument document) throws InvalidFormatException, IOException {
		stylingDoc.HeaderFooter(document);// Styling of Header and footer

		stylingDoc.bold(document);
		FileOutputStream out = new FileOutputStream(word + passData.Exceldata("Account Name") + fileName + ".docx");

		document.write(out);
		out.close();
		logger.info("Doc file is Ready for " + passData.Exceldata("Account Name"));

	}

	public static void stats(XWPFDocument document)
			throws NumberFormatException, InvalidFormatException, FileNotFoundException, IOException, XmlException {
		XWPFParagraph Statsparagraph = document.createParagraph();
		XWPFRun run = Statsparagraph.createRun();
		Statsparagraph = document.createParagraph();
		run = Statsparagraph.createRun();
		stylingDoc.setNoProof(run);
		stylingDoc.FontFamilySize(run);
		logger.info("Start writing in Doc file");
		logger.info("Printing workspaces");
		try {
			run.setText(passData.Exceldata("Account Name") + "\nhas\n" + passData.Exceldata("Total WorkSpace")
					+ "\nworkspaces. Stats below are a sum of assets found across all workspaces.");
		} catch (Exception e) {
			// TODO: handle exception
			run.setText(passData.Exceldata("Account Name") + "\nhas\n" + passData.Exceldata("Total WorkSpace")
					+ "\nworkspaces. Stats below are a sum of assets found across all workspaces.");
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
//				XWPFRun model = paragraphModel.createRun();
//				model.setBold(true);
//				model.addCarriageReturn();
//				model.setText("Snippets");

				XWPFParagraph SnippetsData = document.createParagraph();
				XWPFRun SnippetsRun = SnippetsData.createRun();
				stylingDoc.setNoProof(SnippetsRun);
				stylingDoc.FontFamilySize(SnippetsRun);
				SnippetsRun.setText(passData.Exceldata("Account Name") + passData.No_Snippets);

				// XWPFParagraph paragraphModel2 = document.createParagraph();
				XWPFRun modellink = SnippetsData.createRun();
				stylingDoc.setNoProof(modellink);
				modellink.setText(
						"\nhttps://experienceleague.adobe.com/docs/marketo/using/product-docs/personalization/segmentation-and-snippets/snippets/create-a-snippet.html?lang=en");
				paragraphModel.createHyperlinkRun(
						"https://experienceleague.adobe.com/docs/marketo/using/product-docs/personalization/segmentation-and-snippets/snippets/create-a-snippet.html?lang=en");
				modellink.setUnderline(UnderlinePatterns.SINGLE);
				modellink.setColor("3333cc");
				stylingDoc.FontFamilySize(modellink);
				logger.info("Stats Section completed");

			}
		} catch (Exception e) {

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

			for (int i = 1; i < 3; i++) {

				img.addCarriageReturn();
				img.addPicture(new FileInputStream(passData.FetchScreenshot("Tags" + i)), Document.PICTURE_TYPE_PNG,
						passData.FetchScreenshot("Tags" + i), Units.toEMU(400), Units.toEMU(210));
				img.addCarriageReturn();
			}

		}

		catch (Exception e) {
			logger.warn("Test is failed hence some data is missing from Program Section\n");

		}

	}

	public static void models(XWPFDocument document)
			throws InvalidFormatException, FileNotFoundException, IOException, XmlException {
		// TODO Auto-generated method stub

		XWPFParagraph paragraphModel = document.createParagraph();
		XWPFRun model = paragraphModel.createRun();
		model.setBold(true);

		model.addCarriageReturn();
		model.addCarriageReturn();
		logger.info("Printing Models Data");
		model.setText("Models");

		try {

			int Model = Integer.parseInt(passData.Exceldata("Models"));
			if (Model > 0) {
				logger.info("Models data is greater then 0 hence printing appropriate data");
				XWPFParagraph paragraphModelData = document.createParagraph();
				XWPFRun modelData = paragraphModelData.createRun();
				stylingDoc.setNoProof(modelData);
				stylingDoc.FontFamilySize(modelData);
				try {
					modelData.setText(passData.Exceldata("Account Name") + ",\nHas previously built\n"
							+ passData.Exceldata("Models") + "\n models in Marketo. Revenue cycle models take marketing"
							+ " to the next level. They model all the stages of your entire revenue funnel—from when you first interact with a lead "
							+ "all the way until the lead is a won customer.");
					modelData = paragraphModelData.createRun();
					modelData.addBreak();
					stylingDoc.setNoProof(modelData);
					stylingDoc.FontFamilySize(modelData);
					modelData.setText(passData.Exceldata("Account Name") + passData.models2);
				} catch (Exception e) {
					// TODO: handle exception
				}

				int j = 1;
				int WPCount = Integer.parseInt(passData.Exceldata("Total WorkSpace"));

				for (int i = 1; i <= WPCount; i++) {
					String data = "Approved Models" + j++;

					try {

						modelData = paragraphModelData.createRun();
						modelData.addBreak();
						modelData.setBold(true);
						modelData.setText("Workspace: " + passData.Exceldata(data, 2));

						XWPFParagraph model_img = document.createParagraph();
						logger.info("Printing images in doc");
						XWPFRun img1 = model_img.createRun();

						img1.addPicture(new FileInputStream(passData.FetchScreenshot(passData.Exceldata(data, 2))),
								Document.PICTURE_TYPE_PNG,
								passData.FetchScreenshot(passData.Exceldata(data, 2)).toString(), Units.toEMU(250),
								Units.toEMU(130));

					}

					catch (Exception e) {
						// TODO: handle exception
						logger.warn("Test is failed to retrive models images\n");
						continue;
					}

					int ModelsCount = Integer.parseInt(passData.Exceldata(data));
					ModelsCount = ModelsCount + 3;
					logger.info(ModelsCount);

					if (ModelsCount > 0) {

						XWPFParagraph model_i = document.createParagraph();
						logger.info("Printing images in doc");
						XWPFRun img = model_i.createRun();
						img.addBreak();
						img.setFontSize(11);
						img.setBold(true);
						img.setText("");

						for (int k = 3; k <= ModelsCount; k++) {
							img.setText(passData.Exceldata(data, k));
							img.addBreak();
							img.addPicture(
									new FileInputStream(
											passData.FetchScreenshotForApprovedModels(passData.Exceldata(data, k))),
									Document.PICTURE_TYPE_PNG,
									passData.FetchScreenshotForApprovedModels(passData.Exceldata(data, k)).toString(),
									Units.toEMU(250), Units.toEMU(130));
							img.addCarriageReturn();
						}
					}
				}
			}

			else {
				logger.info("you hadn't added any models to your subscription " + passData.Exceldata("Account Name"));
				XWPFParagraph paragraphModelData = document.createParagraph();

				XWPFRun modelTest = paragraphModelData.createRun();
				stylingDoc.setNoProof(modelTest);
				modelTest.setFontSize(11);
				stylingDoc.setNoProof(modelTest);
				stylingDoc.FontFamilySize(modelTest);
				modelTest.setText(passData.Exceldata("Account Name") + passData.No_models);

				XWPFRun modelTest1 = paragraphModelData.createRun();
				modelTest1.addCarriageReturn();
				modelTest1.addCarriageReturn();
				stylingDoc.setNoProof(modelTest1);
				stylingDoc.FontFamilySize(modelTest1);
				modelTest1.setText("Marketo Docs: Create a New Revenue Model -");

				XWPFRun modelTestLink1 = paragraphModelData.createRun();
				stylingDoc.FontFamilySize(modelTestLink1);
				stylingDoc.setNoProof(modelTestLink1);
				modelTestLink1.setColor("3333cc");
				stylingDoc.setNoProof(modelTest1);
				modelTestLink1.setUnderline(UnderlinePatterns.SINGLE);
				// paragraphModelData.createHyperlinkRun("https://experienceleague.adobe.com/docs/marketo/using/product-docs/reporting/revenue-cycle-analytics/revenue-cycle-models/create-a-new-revenue-model.html");
				modelTestLink1.setText(
						"https://experienceleague.adobe.com/docs/marketo/using/product-docs/reporting/revenue-cycle-analytics/revenue-cycle-models/create-a-new-revenue-model.html");

				XWPFRun modelTest2 = paragraphModelData.createRun();
				modelTest2.addCarriageReturn();
				stylingDoc.setNoProof(modelTest2);
				stylingDoc.FontFamilySize(modelTest2);
				modelTest2.setText("Marketo Docs: Understanding Revenue Models-");

				XWPFRun modelTestLink2 = paragraphModelData.createRun();
				stylingDoc.FontFamilySize(modelTestLink2);
				modelTestLink2.setColor("3333cc");
				stylingDoc.setNoProof(modelTestLink2);
				modelTestLink2.setUnderline(UnderlinePatterns.SINGLE);
				modelTestLink2.setText(
						"https://experienceleague.adobe.com/docs/marketo/using/product-docs/reporting/revenue-cycle-analytics/revenue-cycle-models/understanding-revenue-models.html ");
				// paragraphModelData.createHyperlinkRun("https://experienceleague.adobe.com/docs/marketo/using/product-docs/reporting/revenue-cycle-analytics/revenue-cycle-models/understanding-revenue-models.html
				// ");

				XWPFRun modelTest3 = paragraphModelData.createRun();
				modelTest3.addCarriageReturn();
				stylingDoc.setNoProof(modelTest3);
				stylingDoc.FontFamilySize(modelTest3);
				modelTest3.setText("Marketo Docs: Understanding Revenue Models-");

				XWPFRun modelTestLink3 = paragraphModelData.createRun();
				stylingDoc.FontFamilySize(modelTestLink3);
				modelTestLink3.setColor("3333cc");
				stylingDoc.setNoProof(modelTestLink3);
				modelTestLink3.setUnderline(UnderlinePatterns.SINGLE);
				// paragraphModelData.createHyperlinkRun("https://experienceleague.adobe.com/docs/marketo/using/product-docs/reporting/revenue-cycle-analytics/revenue-cycle-models/understanding-revenue-models.html
				// ");
				modelTestLink3.setText(
						"https://nation.marketo.com/t5/product-blogs/marketo-revenue-attribution-explained/ba-p/244033");

				XWPFRun modelTest4 = paragraphModelData.createRun();
				modelTest4.addCarriageReturn();

				stylingDoc.setNoProof(modelTest4);
				stylingDoc.FontFamilySize(modelTest4);
				modelTest4.setText("Marketo-Fu - Episode 15: Attribution Basics -");

				XWPFRun modelTestLink4 = paragraphModelData.createRun();
				stylingDoc.FontFamilySize(modelTestLink3);
				modelTestLink4.setColor("3333cc");
				stylingDoc.setNoProof(modelTestLink4);
				modelTestLink4.setUnderline(UnderlinePatterns.SINGLE);
				// paragraphModelData.createHyperlinkRun("https://experienceleague.adobe.com/docs/marketo/using/product-docs/reporting/revenue-cycle-analytics/revenue-cycle-models/understanding-revenue-models.html
				// ");
				modelTestLink4.setText("https://youtu.be/Oy_Zqdu3SeI");

				logger.info("Model section completed");

			}

		} catch (

		Exception e) {
			logger.error("Test is failed hence does not found any data it having null values\n");
			// TODO: handle exception
		}

	}

	public static void lead(XWPFDocument document)
			throws NumberFormatException, InvalidFormatException, FileNotFoundException, IOException, XmlException {
		XWPFParagraph LeadParagraph = document.createParagraph();
		XWPFRun LeadRun = LeadParagraph.createRun();
		LeadRun.addCarriageReturn();
		LeadRun.setBold(true);
		logger.info("Printing Lead scoring Section");
		LeadRun.setText("Lead Scoring");
		// adding lead scoring bullet point data into report
		try {
			int ChangeScore = Integer.parseInt(passData.Exceldata("Change Score"));
			if (ChangeScore > 10) {
				XWPFParagraph LeadData = document.createParagraph();
				XWPFRun LeadDataRun = LeadData.createRun();
				stylingDoc.FontFamilySize(LeadDataRun);
				stylingDoc.setNoProof(LeadDataRun);
				LeadData.setSpacingAfter(0);
				LeadData.setNumID(stylingDoc.bullet(document));
				stylingDoc.FontFamilySize(LeadDataRun);
				LeadDataRun.setText(passData.LeadScoring(passData.Exceldata("Change Score")));
				XWPFParagraph LeadDatatoken = document.createParagraph();
				XWPFRun LeadDataRuntoken = LeadDatatoken.createRun();
				stylingDoc.FontFamilySize(LeadDataRuntoken);
				stylingDoc.setNoProof(LeadDataRuntoken);
				LeadDatatoken.setSpacingAfter(0);
				LeadDatatoken.setNumID(stylingDoc.bullet(document));
				stylingDoc.FontFamilySize(LeadDataRuntoken);
				LeadDataRuntoken.setText(passData.no_tokens);
				XWPFParagraph LeadDatatlink = document.createParagraph();
				XWPFRun LeadDataLink = LeadDatatlink.createRun();
				stylingDoc.FontFamilySize(LeadDataLink);
				stylingDoc.setNoProof(LeadDataLink);
				LeadDataLink.setUnderline(UnderlinePatterns.SINGLE);
				LeadDataLink.setColor("3333cc");
				LeadDatatlink.setSpacingAfter(0);
				LeadDataLink.setText("https://docs.marketo.com/display/public/DOCS/Managing+My+Tokens");
				XWPFRun LeadDatabelow = LeadDatatlink.createRun();
				LeadDatabelow.addCarriageReturn();
				LeadDatabelow.addCarriageReturn();
				stylingDoc.FontFamilySize(LeadDatabelow);
				stylingDoc.setNoProof(LeadDatabelow);
				LeadDatabelow.setText(passData.lead_data);
			} else if (ChangeScore <= 10 && ChangeScore >= 0) {
				XWPFParagraph LeadData = document.createParagraph();
				XWPFRun LeadDataRun = LeadData.createRun();
				stylingDoc.FontFamilySize(LeadDataRun);
				stylingDoc.setNoProof(LeadDataRun);
				LeadData.setSpacingAfter(0);
				LeadData.setNumID(stylingDoc.bullet(document));
				stylingDoc.FontFamilySize(LeadDataRun);
				LeadDataRun.setText(
						passData.Exceldata("Account Name") + "\nhas a total of (0-10) lead scoring campaigns.");
				XWPFParagraph LeadDatatoken = document.createParagraph();
				XWPFRun LeadDataRuntoken = LeadDatatoken.createRun();
				stylingDoc.FontFamilySize(LeadDataRuntoken);
				stylingDoc.setNoProof(LeadDataRuntoken);
				LeadDatatoken.setSpacingAfter(0);
				LeadDatatoken.setNumID(stylingDoc.bullet(document));
				stylingDoc.FontFamilySize(LeadDataRuntoken);
				LeadDataRuntoken.setText(passData.no_tokens);
				XWPFParagraph LeadDatatlink = document.createParagraph();
				XWPFRun LeadDataLink = LeadDatatlink.createRun();
				stylingDoc.FontFamilySize(LeadDataLink);
				stylingDoc.setNoProof(LeadDataLink);
				LeadDataLink.setUnderline(UnderlinePatterns.SINGLE);
				LeadDataLink.setColor("3333cc");
				LeadDatatlink.setSpacingAfter(0);
				LeadDataLink.setText("https://docs.marketo.com/display/public/DOCS/Managing+My+Tokens");
				XWPFRun LeadDatapara1 = LeadData.createRun();
				LeadDatapara1.addCarriageReturn();
				LeadDatapara1.addCarriageReturn();
				stylingDoc.FontFamilySize(LeadDatapara1);
				stylingDoc.setNoProof(LeadDatapara1);
				LeadDatapara1.setText(passData.No_lead_scoring);
				XWPFRun LeadDatapara2 = LeadData.createRun();
				LeadDatapara2.addCarriageReturn();
				LeadDatapara2.addCarriageReturn();
				stylingDoc.FontFamilySize(LeadDatapara2);
				stylingDoc.setNoProof(LeadDatapara2);
				LeadDatapara2.setText(passData.No_lead_scoring2);
				XWPFRun LeadDatapara3 = LeadData.createRun();
				stylingDoc.FontFamilySize(LeadDatapara3);
				LeadDatapara3.addCarriageReturn();
				LeadDatapara3.addCarriageReturn();
				LeadDatapara3.setText("I would strongly encourage you to share the following lead scoring resources:");
				LeadDatapara3.addCarriageReturn();
				LeadDatapara3.setText("Marketo Resources: Lead Scoing:");
				XWPFRun LeadDataparaLink3 = LeadData.createRun();
				stylingDoc.FontFamilySize(LeadDataparaLink3);
				stylingDoc.setNoProof(LeadDataparaLink3);
				LeadDataparaLink3.setUnderline(UnderlinePatterns.SINGLE);
				LeadDataparaLink3.setColor("3333cc");
				LeadDataparaLink3.setText("\nhttps://www.marketo.com/resources/lead-scoring/");
				XWPFRun LeadDatapara4 = LeadData.createRun();
				stylingDoc.FontFamilySize(LeadDatapara4);
				LeadDatapara4.addCarriageReturn();
				LeadDatapara4.setText("Marketo's Definitive Guide to Lead Scoring");
				XWPFRun LeadDataparaLink4 = LeadData.createRun();
				stylingDoc.FontFamilySize(LeadDataparaLink4);
				stylingDoc.setNoProof(LeadDataparaLink4);
				LeadDataparaLink4.setUnderline(UnderlinePatterns.SINGLE);
				LeadDataparaLink4.setColor("3333cc");
				LeadDataparaLink4.setText("\nhttps://www.marketo.com/definitive-guides/lead-scoring/");
				XWPFRun LeadDatapara5 = LeadData.createRun();
				stylingDoc.FontFamilySize(LeadDatapara5);
				LeadDatapara5.addCarriageReturn();
				LeadDatapara5.setText("Lead Scoring Cheat Sheet:");
				XWPFRun LeadDataparaLink5 = LeadData.createRun();
				stylingDoc.FontFamilySize(LeadDataparaLink5);
				stylingDoc.setNoProof(LeadDataparaLink5);
				LeadDataparaLink5.setUnderline(UnderlinePatterns.SINGLE);
				LeadDataparaLink5.setColor("3333cc");
				LeadDataparaLink5.setText("\nhttps://www.marketo.com/cheat-sheets/lead-scoring/");
				XWPFRun LeadDatapara6 = LeadData.createRun();
				stylingDoc.FontFamilySize(LeadDatapara6);
				LeadDatapara6.addCarriageReturn();
				LeadDatapara6.setText("Marketo Docs: Simple Scoring:");
				XWPFRun LeadDataparaLink6 = LeadData.createRun();
				stylingDoc.FontFamilySize(LeadDataparaLink6);
				stylingDoc.setNoProof(LeadDataparaLink6);
				LeadDataparaLink6.setUnderline(UnderlinePatterns.SINGLE);
				LeadDataparaLink6.setColor("3333cc");
				LeadDataparaLink6.setText(
						"\nhttps://experienceleague.adobe.com/docs/marketo/using/getting-started-with-marketo/quick-wins/simple-scoring.html?lang=en");
			}
		} catch (Exception e) {
			// TODO: handle exception
		}
	}

	public static void intrestingMoment(XWPFDocument document)
			throws NumberFormatException, InvalidFormatException, FileNotFoundException, IOException {

		if (passData.Exceldata("Interesting Moment Subscription").equalsIgnoreCase("true")) {
			logger.info("User have Interesting Moment Subscription");
			XWPFParagraph Interesting = document.createParagraph();
			XWPFRun InterestingRun = Interesting.createRun();
			InterestingRun.setBold(true);
			InterestingRun.addCarriageReturn();
			InterestingRun.addCarriageReturn();
			InterestingRun.setText("Interesting Moment");

			try {
				int Interesting_Moment = Integer.parseInt(passData.Exceldata("Interesting Moment"));

				if (Interesting_Moment > 5) {

					XWPFParagraph InterestingData = document.createParagraph();
					XWPFRun InterestingDatarun = InterestingData.createRun();

					stylingDoc.setNoProof(InterestingDatarun);
					stylingDoc.FontFamilySize(InterestingDatarun);
					try {
						InterestingDatarun.setText("Client has " + passData.Exceldata("Interesting Moment")
								+ "\nInteresting Moments"
								+ ".\nThe following screenshot shows some Interesting Moments that have been defined by the client.\n"
								+ passData.intresting_moment);
					} catch (Exception e) {
						// TODO: handle exception
					}

					InterestingDatarun = InterestingData.createRun();
					for (int i = 1; i < 4; i++) {
						try {
							InterestingDatarun.addCarriageReturn();
							InterestingDatarun.addPicture(
									new FileInputStream(passData.FetchScreenshot("Interesting Moment" + i)),
									Document.PICTURE_TYPE_PNG, passData.FetchScreenshot("Interesting Moment"),
									Units.toEMU(470), Units.toEMU(80));
							InterestingDatarun.addCarriageReturn();
						} catch (Exception e) {

							// TODO: handle exception
						}
					}

					InterestingData = document.createParagraph();
					InterestingDatarun = InterestingData.createRun();

					stylingDoc.setNoProof(InterestingDatarun);
					stylingDoc.FontFamilySize(InterestingDatarun);
					InterestingDatarun.setText(passData.intresting_moment_below);
				}

				else if (Interesting_Moment <= 5 && Interesting_Moment > 0) {
					XWPFParagraph InterestingData = document.createParagraph();
					XWPFRun InterestingDatarun = InterestingData.createRun();

					stylingDoc.setNoProof(InterestingDatarun);
					stylingDoc.FontFamilySize(InterestingDatarun);
					try {
						InterestingDatarun
								.setText(passData.Exceldata("Account Name") + passData.intresting_moment_less_5);

					} catch (Exception e) {
						// TODO: handle exception
					}

					InterestingDatarun = InterestingData.createRun();
					for (int i = 1; i < 4; i++) {
						try {
							InterestingDatarun.addCarriageReturn();
							InterestingDatarun.addPicture(
									new FileInputStream(passData.FetchScreenshot("Interesting Moment" + i)),
									Document.PICTURE_TYPE_PNG, passData.FetchScreenshot("Interesting Moment"),
									Units.toEMU(470), Units.toEMU(80));

						} catch (Exception e) {

							// TODO: handle exception
						}
					}

					XWPFRun InterestingMomentdata = InterestingData.createRun();
					stylingDoc.FontFamilySize(InterestingMomentdata);
					InterestingMomentdata.setText(passData.intresting_momentData);

					XWPFRun InterestingMoment1 = InterestingData.createRun();
					stylingDoc.FontFamilySize(InterestingMoment1);
					InterestingMoment1.addCarriageReturn();
					InterestingMoment1.addCarriageReturn();
					InterestingMoment1
							.setText("I would encourage you to share the following interesting moment’s resources:");
					InterestingMoment1.addCarriageReturn();
					InterestingMoment1.setText("Marketo Docs: Interesting Moment Overview: ");

					XWPFRun InterestingMomentLink = InterestingData.createRun();
					stylingDoc.FontFamilySize(InterestingMomentLink);
					stylingDoc.setNoProof(InterestingMomentLink);
					InterestingMomentLink.setUnderline(UnderlinePatterns.SINGLE);
					InterestingMomentLink.setColor("3333cc");
					InterestingMomentLink.setText(
							"\nhttps://experienceleague.adobe.com/docs/marketo/using/product-docs/core-marketo-concepts/smart-campaigns/flow-actions/interesting-moment.html");

					XWPFRun InterestingMoment2 = InterestingData.createRun();
					stylingDoc.FontFamilySize(InterestingMoment2);
					InterestingMoment2.addCarriageReturn();
					InterestingMoment2.setText("Marketo Docs: Using Interesting Moments:");

					// InterestingMoment2.setText("\nUnderstanding Event Programs:");

					XWPFRun InterestingMomentLink2 = InterestingData.createRun();
					stylingDoc.FontFamilySize(InterestingMomentLink2);
					stylingDoc.setNoProof(InterestingMomentLink2);
					InterestingMomentLink2.setUnderline(UnderlinePatterns.SINGLE);
					InterestingMomentLink2.setColor("3333cc");
					InterestingMomentLink2.setText(
							"\nhttps://experienceleague.adobe.com/docs/marketo/using/product-docs/marketo-sales-insight/msi-for-salesforce/features/tabs-in-the-msi-panel/interesting-moments/using-interesting-moments.html?lang=en");

					XWPFRun InterestingMoment3 = InterestingData.createRun();
					stylingDoc.FontFamilySize(InterestingMoment3);
					InterestingMoment3.addCarriageReturn();
					InterestingMoment3.setText("How to Create a Custom Interesting Moment Type:");

					// InterestingMoment2.setText("\nUnderstanding Event Programs:");

					XWPFRun InterestingMomentLink3 = InterestingData.createRun();
					stylingDoc.FontFamilySize(InterestingMomentLink3);
					stylingDoc.setNoProof(InterestingMomentLink3);
					InterestingMomentLink3.setUnderline(UnderlinePatterns.SINGLE);
					InterestingMomentLink3.setColor("3333cc");
					InterestingMomentLink3.setText(
							"\nhttps://nation.marketo.com/t5/knowledgebase/how-to-create-a-custom-interesting-moment-type/ta-p/253526");

					XWPFRun InterestingMoment4 = InterestingData.createRun();
					stylingDoc.FontFamilySize(InterestingMoment4);
					InterestingMoment4.addCarriageReturn();
					InterestingMoment4.setText("Using Interesting Moments Best Practices – DemandSpring: ");

					// InterestingMoment2.setText("\nUnderstanding Event Programs:");

					XWPFRun InterestingMomentLink4 = InterestingData.createRun();
					stylingDoc.FontFamilySize(InterestingMomentLink4);
					stylingDoc.setNoProof(InterestingMomentLink4);
					InterestingMomentLink4.setUnderline(UnderlinePatterns.SINGLE);
					InterestingMomentLink4.setColor("3333cc");
					InterestingMomentLink4
							.setText("\nhttps://demandspring.com/blog/using-interesting-moments-best-practices/");

					XWPFRun InterestingMoment5 = InterestingData.createRun();
					stylingDoc.FontFamilySize(InterestingMoment5);
					InterestingMoment5.addCarriageReturn();
					InterestingMoment5.setText("Marketo Interesting Moments vs. Tasks – MarketingRockStarGuides:");

					// InterestingMoment2.setText("\nUnderstanding Event Programs:");

					XWPFRun InterestingMomentLink5 = InterestingData.createRun();
					stylingDoc.FontFamilySize(InterestingMomentLink5);
					stylingDoc.setNoProof(InterestingMomentLink5);
					InterestingMomentLink5.setUnderline(UnderlinePatterns.SINGLE);
					InterestingMomentLink5.setColor("3333cc");
					InterestingMomentLink5.setText(
							"\nhttps://www.marketingrockstarguides.com/marketo-interesting-moments-vs-tasks-1206/");

				} else {

					XWPFParagraph InterestingData = document.createParagraph();
					XWPFRun InterestingDatarun = InterestingData.createRun();
					stylingDoc.FontFamilySize(InterestingDatarun);
					stylingDoc.setNoProof(InterestingDatarun);
					InterestingDatarun.setText(passData.intresting_moment_else);

					XWPFRun InterestingMomentdata = InterestingData.createRun();
					stylingDoc.FontFamilySize(InterestingMomentdata);
					InterestingMomentdata.setText(passData.intresting_momentData);

					XWPFRun InterestingMoment1 = InterestingData.createRun();
					stylingDoc.FontFamilySize(InterestingMoment1);
					InterestingMoment1.addCarriageReturn();
					InterestingMoment1.addCarriageReturn();
					InterestingMoment1
							.setText("I would encourage you to share the following interesting moment’s resources:");
					InterestingMoment1.addCarriageReturn();
					InterestingMoment1.setText("Marketo Docs: Interesting Moment Overview: ");

					XWPFRun InterestingMomentLink = InterestingData.createRun();
					stylingDoc.FontFamilySize(InterestingMomentLink);
					stylingDoc.setNoProof(InterestingMomentLink);
					InterestingMomentLink.setUnderline(UnderlinePatterns.SINGLE);
					InterestingMomentLink.setColor("3333cc");
					InterestingMomentLink.setText(
							"\nhttps://experienceleague.adobe.com/docs/marketo/using/product-docs/core-marketo-concepts/smart-campaigns/flow-actions/interesting-moment.html");

					XWPFRun InterestingMoment2 = InterestingData.createRun();
					stylingDoc.FontFamilySize(InterestingMoment2);
					InterestingMoment2.addCarriageReturn();
					InterestingMoment2.setText("Marketo Docs: Using Interesting Moments:");

					// InterestingMoment2.setText("\nUnderstanding Event Programs:");

					XWPFRun InterestingMomentLink2 = InterestingData.createRun();
					stylingDoc.FontFamilySize(InterestingMomentLink2);
					stylingDoc.setNoProof(InterestingMomentLink2);
					InterestingMomentLink2.setUnderline(UnderlinePatterns.SINGLE);
					InterestingMomentLink2.setColor("3333cc");
					InterestingMomentLink2.setText(
							"\nhttps://experienceleague.adobe.com/docs/marketo/using/product-docs/marketo-sales-insight/msi-for-salesforce/features/tabs-in-the-msi-panel/interesting-moments/using-interesting-moments.html?lang=en");

					XWPFRun InterestingMoment3 = InterestingData.createRun();
					stylingDoc.FontFamilySize(InterestingMoment3);
					InterestingMoment3.addCarriageReturn();
					InterestingMoment3.setText("How to Create a Custom Interesting Moment Type:");

					// InterestingMoment2.setText("\nUnderstanding Event Programs:");

					XWPFRun InterestingMomentLink3 = InterestingData.createRun();
					stylingDoc.FontFamilySize(InterestingMomentLink3);
					stylingDoc.setNoProof(InterestingMomentLink3);
					InterestingMomentLink3.setUnderline(UnderlinePatterns.SINGLE);
					InterestingMomentLink3.setColor("3333cc");
					InterestingMomentLink3.setText(
							"\nhttps://nation.marketo.com/t5/knowledgebase/how-to-create-a-custom-interesting-moment-type/ta-p/253526");

					XWPFRun InterestingMoment4 = InterestingData.createRun();
					stylingDoc.FontFamilySize(InterestingMoment4);
					InterestingMoment4.addCarriageReturn();
					InterestingMoment4.setText("Using Interesting Moments Best Practices – DemandSpring: ");

					// InterestingMoment2.setText("\nUnderstanding Event Programs:");

					XWPFRun InterestingMomentLink4 = InterestingData.createRun();
					stylingDoc.FontFamilySize(InterestingMomentLink4);
					stylingDoc.setNoProof(InterestingMomentLink4);
					InterestingMomentLink4.setUnderline(UnderlinePatterns.SINGLE);
					InterestingMomentLink4.setColor("3333cc");
					InterestingMomentLink4
							.setText("\nhttps://demandspring.com/blog/using-interesting-moments-best-practices/");

					XWPFRun InterestingMoment5 = InterestingData.createRun();
					stylingDoc.FontFamilySize(InterestingMoment5);
					InterestingMoment5.addCarriageReturn();
					InterestingMoment5.setText("Marketo Interesting Moments vs. Tasks – MarketingRockStarGuides:");

					// InterestingMoment2.setText("\nUnderstanding Event Programs:");

					XWPFRun InterestingMomentLink5 = InterestingData.createRun();
					stylingDoc.FontFamilySize(InterestingMomentLink5);
					stylingDoc.setNoProof(InterestingMomentLink5);
					InterestingMomentLink5.setUnderline(UnderlinePatterns.SINGLE);
					InterestingMomentLink5.setColor("3333cc");
					InterestingMomentLink5.setText(
							"\nhttps://www.marketingrockstarguides.com/marketo-interesting-moments-vs-tasks-1206/");

				}

			}

			catch (NumberFormatException ex) { // handle your exception
				logger.info("Test is skipped or fail hence data is missing for Intresting Moment\n");
			}
			logger.info("Interesting Moment part is done");
		}
	}

	public static void DataManagment(XWPFDocument document) throws IOException {
		XWPFParagraph DataManagement = document.createParagraph();
		XWPFRun DataManagementRun = DataManagement.createRun();
		DataManagementRun.setBold(true);
		DataManagementRun.addCarriageReturn();
		DataManagementRun.addCarriageReturn();
		logger.info("Printing data of Data Management...");
		DataManagementRun.setText("Data Management");
		try {
			int DataManagment = Integer.parseInt(passData.Exceldata("Change Data Value"));
			if (DataManagment > 0) {
				XWPFParagraph DataManagementData = document.createParagraph();
				XWPFRun DataManagementDatarun = DataManagementData.createRun();
				stylingDoc.FontFamilySize(DataManagementDatarun);
				stylingDoc.setNoProof(DataManagementDatarun);
				DataManagementDatarun.setText(passData.Exceldata("Account Name") + passData.Data);
			} else {
				XWPFParagraph DataManagementData = document.createParagraph();
				XWPFRun DataManagementDatarun = DataManagementData.createRun();
				stylingDoc.FontFamilySize(DataManagementDatarun);
				stylingDoc.setNoProof(DataManagementDatarun);
				DataManagementDatarun.setText(passData.No_Data);
				XWPFRun DataManagementDatarun2 = DataManagementData.createRun();
				stylingDoc.FontFamilySize(DataManagementDatarun2);
				DataManagementDatarun2.addCarriageReturn();
				DataManagementDatarun2.addCarriageReturn();
				stylingDoc.setNoProof(DataManagementDatarun2);
				DataManagementDatarun2.setText(passData.No_data_below);
				XWPFRun DataManagementLink = DataManagementData.createRun();
				stylingDoc.FontFamilySize(DataManagementLink);
				stylingDoc.setNoProof(DataManagementLink);
				DataManagementLink.setUnderline(UnderlinePatterns.SINGLE);
				DataManagementLink.setColor("3333cc");
				DataManagementLink.setText(
						"\nhttps://experienceleague.adobe.com/docs/marketo/using/product-docs/core-marketo-concepts/smart-lists-and-static-lists/managing-people-in-smart-lists/add-person-to-blocklist.html?lang=en  ");
				XWPFRun DataManagementDatarun3 = DataManagementData.createRun();
				stylingDoc.FontFamilySize(DataManagementDatarun3);
				DataManagementDatarun3.addCarriageReturn();
				stylingDoc.setNoProof(DataManagementDatarun3);
				DataManagementDatarun3
						.setText("Here is a high overview on how to create Change Data Value flow actions:");
				XWPFRun DataManagementLink1 = DataManagementData.createRun();
				stylingDoc.FontFamilySize(DataManagementLink1);
				stylingDoc.setNoProof(DataManagementLink1);
				DataManagementLink1.setUnderline(UnderlinePatterns.SINGLE);
				DataManagementLink1.setColor("3333cc");
				DataManagementLink1.setText("\nhttps://docs.marketo.com/display/public/DOCS/Change+Data+Value");
			}
		} catch (Exception e) {
			// TODO: handle exception
		}
	}

	public static void Events(XWPFDocument document) throws IOException {
		XWPFParagraph Events = document.createParagraph();
		XWPFRun Eventsrun = Events.createRun();
		Eventsrun.setBold(true);
		Eventsrun.addCarriageReturn();
		logger.info("Event Data is printing");
		Eventsrun.setText("Events");
		try {
			int Event_Program = Integer.parseInt(passData.Exceldata("Event Programs"));
			if (Event_Program >= 5) {
				XWPFParagraph EventsData = document.createParagraph();
				XWPFRun EventsDatarun = EventsData.createRun();
				stylingDoc.FontFamilySize(EventsDatarun);
				stylingDoc.setNoProof(EventsDatarun);
				EventsDatarun.setText(passData.Exceldata("Account Name") + "\nhas\n"
						+ passData.Exceldata("Event Programs") + "\nEvent campaigns in Marketo.\n" + passData.events);

				XWPFParagraph imgPara = document.createParagraph();
				XWPFRun img = imgPara.createRun();
				for (int i = 1; i < 2; i++) {

					img.addCarriageReturn();
					img.addPicture(new FileInputStream(passData.FetchScreenshot("Event")), Document.PICTURE_TYPE_PNG,
							passData.FetchScreenshot("Event"), Units.toEMU(200), Units.toEMU(380));
					img.addCarriageReturn();
				}

			}

			else if (Event_Program < 5 && Event_Program > 0) {
				XWPFParagraph EventsData = document.createParagraph();
				XWPFRun EventsDatarun = EventsData.createRun();
				stylingDoc.FontFamilySize(EventsDatarun);
				stylingDoc.setNoProof(EventsDatarun);
				EventsDatarun.setText(passData.Exceldata("Account Name") + passData.event2);

				XWPFParagraph imgPara = document.createParagraph();
				XWPFRun img = imgPara.createRun();
				for (int i = 1; i < 2; i++) {

					img.addCarriageReturn();
					img.addPicture(new FileInputStream(passData.FetchScreenshot("Event")), Document.PICTURE_TYPE_PNG,
							passData.FetchScreenshot("Event"), Units.toEMU(200), Units.toEMU(380));
					img.addCarriageReturn();
				}

				XWPFRun EventsData1 = EventsData.createRun();
				stylingDoc.FontFamilySize(EventsData1);
				EventsData1.addCarriageReturn();
				EventsData1.addCarriageReturn();
				EventsData1.setText("Share the following resources:\r\n");
				EventsData1.addCarriageReturn();
				EventsData1.setText("\nUnderstanding Event Programs:");
				XWPFRun EventsDataLink1 = EventsData.createRun();
				stylingDoc.FontFamilySize(EventsDataLink1);
				stylingDoc.setNoProof(EventsDataLink1);
				EventsDataLink1.setUnderline(UnderlinePatterns.SINGLE);
				EventsDataLink1.setColor("3333cc");
				EventsDataLink1.setText(
						"\nhttps://experienceleague.adobe.com/docs/marketo/using/product-docs/demand-generation/events/understanding-events/understanding-event-programs.html");
				EventsDataLink1.addCarriageReturn();
				XWPFRun EventsData2 = EventsData.createRun();
				stylingDoc.FontFamilySize(EventsData2);
				stylingDoc.setNoProof(EventsData2);
				EventsData2.setText("Create a New Event Program:");
				XWPFRun EventsDataLink2 = EventsData.createRun();
				stylingDoc.FontFamilySize(EventsDataLink2);
				stylingDoc.setNoProof(EventsDataLink2);
				EventsDataLink2.setUnderline(UnderlinePatterns.SINGLE);
				EventsDataLink2.setColor("3333cc");
				EventsDataLink2.setText(
						"\nhttps://experienceleague.adobe.com/docs/marketo/using/product-docs/email-marketing/drip-nurturing/creating-an-engagement-program/create-an-engagement-program.html?lang=en");
				XWPFRun EventsData3 = EventsData.createRun();
				stylingDoc.FontFamilySize(EventsData3);
				stylingDoc.setNoProof(EventsData3);
				EventsData3.addCarriageReturn();
				EventsData3.setText("Edit an Event Channel:");
				XWPFRun EventsDataLink3 = EventsData.createRun();
				stylingDoc.FontFamilySize(EventsDataLink3);
				stylingDoc.setNoProof(EventsDataLink3);
				EventsDataLink3.setUnderline(UnderlinePatterns.SINGLE);
				EventsDataLink3.setColor("3333cc");
				EventsDataLink3.setText(
						"\nhttps://experienceleague.adobe.com/docs/marketo/using/product-docs/demand-generation/events/understanding-events/edit-an-event-channel.html");
				XWPFRun EventsData4 = EventsData.createRun();
				stylingDoc.FontFamilySize(EventsData4);
				stylingDoc.setNoProof(EventsData4);
				EventsData4.addCarriageReturn();
				EventsData4.setText("Adding Members to an Event Program:");
				XWPFRun EventsDataLink4 = EventsData.createRun();
				stylingDoc.FontFamilySize(EventsDataLink4);
				stylingDoc.setNoProof(EventsDataLink4);
				EventsDataLink4.setUnderline(UnderlinePatterns.SINGLE);
				EventsDataLink4.setColor("3333cc");
				EventsDataLink4.setText(
						"\nhttps://experienceleague.adobe.com/docs/marketo/using/product-docs/demand-generation/events/understanding-events/adding-members-to-an-event-program.html?src=contextnavpagetreemode");
				XWPFRun EventsData5 = EventsData.createRun();
				stylingDoc.FontFamilySize(EventsData5);
				stylingDoc.setNoProof(EventsData5);
				EventsData5.addCarriageReturn();
				EventsData5.setText("LaunchPoint Event Partners:");
				XWPFRun EventsDataLink5 = EventsData.createRun();
				stylingDoc.FontFamilySize(EventsDataLink5);
				stylingDoc.setNoProof(EventsDataLink5);
				EventsDataLink5.setUnderline(UnderlinePatterns.SINGLE);
				EventsDataLink5.setColor("3333cc");
				EventsDataLink5.setText(
						"\n: https://experienceleague.adobe.com/docs/marketo/using/product-docs/demand-generation/events/understanding-events/event-partners.html?src=contextnavpagetreemode");
			} else {
				XWPFParagraph EventsData = document.createParagraph();
				XWPFRun EventsDatarun = EventsData.createRun();
				stylingDoc.FontFamilySize(EventsDatarun);
				stylingDoc.setNoProof(EventsDatarun);
				EventsDatarun.setText(passData.No_events);
				XWPFRun EventsData1 = EventsData.createRun();
				stylingDoc.FontFamilySize(EventsData1);
				EventsData1.addCarriageReturn();
				EventsData1.addCarriageReturn();
				EventsData1.setText("Share the following resources:\r\n");
				EventsData1.addCarriageReturn();
				EventsData1.setText("\nUnderstanding Event Programs:");
				XWPFRun EventsDataLink1 = EventsData.createRun();
				stylingDoc.FontFamilySize(EventsDataLink1);
				stylingDoc.setNoProof(EventsDataLink1);
				EventsDataLink1.setUnderline(UnderlinePatterns.SINGLE);
				EventsDataLink1.setColor("3333cc");
				EventsDataLink1.setText(
						"\nhttps://experienceleague.adobe.com/docs/marketo/using/product-docs/demand-generation/events/understanding-events/understanding-event-programs.html");
				EventsDataLink1.addCarriageReturn();
				XWPFRun EventsData2 = EventsData.createRun();
				stylingDoc.FontFamilySize(EventsData2);
				stylingDoc.setNoProof(EventsData2);
				EventsData2.setText("Create a New Event Program:");
				XWPFRun EventsDataLink2 = EventsData.createRun();
				stylingDoc.FontFamilySize(EventsDataLink2);
				stylingDoc.setNoProof(EventsDataLink2);
				EventsDataLink2.setUnderline(UnderlinePatterns.SINGLE);
				EventsDataLink2.setColor("3333cc");
				EventsDataLink2.setText(
						"\nhttps://experienceleague.adobe.com/docs/marketo/using/product-docs/email-marketing/drip-nurturing/creating-an-engagement-program/create-an-engagement-program.html?lang=en");
				XWPFRun EventsData3 = EventsData.createRun();
				stylingDoc.FontFamilySize(EventsData3);
				stylingDoc.setNoProof(EventsData3);
				EventsData3.addCarriageReturn();
				EventsData3.setText("Edit an Event Channel:");
				XWPFRun EventsDataLink3 = EventsData.createRun();
				stylingDoc.FontFamilySize(EventsDataLink3);
				stylingDoc.setNoProof(EventsDataLink3);
				EventsDataLink3.setUnderline(UnderlinePatterns.SINGLE);
				EventsDataLink3.setColor("3333cc");
				EventsDataLink3.setText(
						"\nhttps://experienceleague.adobe.com/docs/marketo/using/product-docs/demand-generation/events/understanding-events/edit-an-event-channel.html");
				XWPFRun EventsData4 = EventsData.createRun();
				stylingDoc.FontFamilySize(EventsData4);
				stylingDoc.setNoProof(EventsData4);
				EventsData4.addCarriageReturn();
				EventsData4.setText("Adding Members to an Event Program:");
				XWPFRun EventsDataLink4 = EventsData.createRun();
				stylingDoc.FontFamilySize(EventsDataLink4);
				stylingDoc.setNoProof(EventsDataLink4);
				EventsDataLink4.setUnderline(UnderlinePatterns.SINGLE);
				EventsDataLink4.setColor("3333cc");
				EventsDataLink4.setText(
						"\nhttps://experienceleague.adobe.com/docs/marketo/using/product-docs/demand-generation/events/understanding-events/adding-members-to-an-event-program.html?src=contextnavpagetreemode");
				XWPFRun EventsData5 = EventsData.createRun();
				stylingDoc.FontFamilySize(EventsData5);
				stylingDoc.setNoProof(EventsData5);
				EventsData5.addCarriageReturn();
				EventsData5.setText("LaunchPoint Event Partners:");
				XWPFRun EventsDataLink5 = EventsData.createRun();
				stylingDoc.FontFamilySize(EventsDataLink5);
				stylingDoc.setNoProof(EventsDataLink5);
				EventsDataLink5.setUnderline(UnderlinePatterns.SINGLE);
				EventsDataLink5.setColor("3333cc");
				EventsDataLink5.setText(
						"\n: https://experienceleague.adobe.com/docs/marketo/using/product-docs/demand-generation/events/understanding-events/event-partners.html?src=contextnavpagetreemode");
			}
		} catch (Exception e) {
			// TODO: handle exception
		}
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

		try {
			int NurtureData_int = Integer.parseInt(passData.Exceldata("Nurture campaigns"));

			if (NurtureData_int > 5)

			{
				NurtureDataRun = NurtureData.createRun();
				stylingDoc.FontFamilySize(NurtureDataRun);
				stylingDoc.setNoProof(NurtureDataRun);
				NurtureDataRun.addCarriageReturn();
				try {
					NurtureDataRun.setText(passData.Exceldata("Account Name") + "\nhas\n"
							+ passData.Exceldata("Nurture campaigns") + passData.Nurture);
				} catch (Exception e) {
					// TODO: handle exception
				}

				XWPFParagraph NurtureData1 = document.createParagraph();
				XWPFRun NurtureDatarun1 = NurtureData1.createRun();
				stylingDoc.FontFamilySize(NurtureDatarun1);

				NurtureData1.setNumID(stylingDoc.bullet(document));

				NurtureDatarun1.setText("Intelligently and automatically deliver content to a target audience.");

				XWPFParagraph NurtureData2 = document.createParagraph();
				XWPFRun NurtureDatarun2 = NurtureData2.createRun();
				stylingDoc.FontFamilySize(NurtureDatarun2);

				NurtureData2.setNumID(stylingDoc.bullet(document));
				NurtureDatarun2.setText(
						"Easily build dialogue with prospects and customers while preventing customers who have already received content from receiving the same content again.");

				XWPFParagraph NurtureData3 = document.createParagraph();
				XWPFRun NurtureDatarun3 = NurtureData3.createRun();
				stylingDoc.FontFamilySize(NurtureDatarun3);
				NurtureData3.setNumID(stylingDoc.bullet(document));
				NurtureDatarun3.setText("Add new content and entire programs to nurture streams.");

				XWPFParagraph NurtureData4 = document.createParagraph();
				XWPFRun NurtureDatarun4 = NurtureData4.createRun();
				stylingDoc.FontFamilySize(NurtureDatarun4);
				NurtureData4.setNumID(stylingDoc.bullet(document));
				NurtureDatarun4.setText("Edit the availability of content.");

				XWPFParagraph NurtureData5 = document.createParagraph();
				XWPFRun NurtureDatarun5 = NurtureData5.createRun();
				stylingDoc.FontFamilySize(NurtureDatarun5);
				NurtureData5.setNumID(stylingDoc.bullet(document));
				NurtureDatarun5
						.setText("Understand content performance based on engagement with each piece of content.");

				XWPFParagraph Nurture_img = document.createParagraph();
				logger.info("Printing images in doc");
				XWPFRun img1 = Nurture_img.createRun();
				for (int i = 1; i < 2; i++) {
					try {

						img1.addPicture(new FileInputStream(passData.FetchScreenshot("Nurture campaigns")),
								Document.PICTURE_TYPE_PNG, passData.FetchScreenshot("Nurture campaigns").toString(),
								Units.toEMU(220), Units.toEMU(380));

					} catch (Exception e) {
						// TODO: handle exception
						logger.warn("Test is failed to retrive models images\n");
					}
				}

			} else if (NurtureData_int <= 5 && NurtureData_int >= 1) {
				NurtureDataRun = NurtureData.createRun();
				stylingDoc.FontFamilySize(NurtureDataRun);
				stylingDoc.setNoProof(NurtureDataRun);
				NurtureDataRun.addCarriageReturn();
				try {
					NurtureDataRun
							.setText(passData.Exceldata("Account Name") + "\nONLY has (1-5)\n" + passData.Nurture2);
				} catch (Exception e) {
					// TODO: handle exception
				}

				XWPFParagraph NurtureData1 = document.createParagraph();
				XWPFRun NurtureDatarun1 = NurtureData1.createRun();
				stylingDoc.FontFamilySize(NurtureDatarun1);

				NurtureData1.setNumID(stylingDoc.bullet(document));

				NurtureDatarun1.setText("Intelligently and automatically deliver content to a target audience.");

				XWPFParagraph NurtureData2 = document.createParagraph();
				XWPFRun NurtureDatarun2 = NurtureData2.createRun();
				stylingDoc.FontFamilySize(NurtureDatarun2);

				NurtureData2.setNumID(stylingDoc.bullet(document));
				NurtureDatarun2.setText(
						"Easily build dialogue with prospects and customers while preventing customers who have already received content from receiving the same content again.");

				XWPFParagraph NurtureData3 = document.createParagraph();
				XWPFRun NurtureDatarun3 = NurtureData3.createRun();
				stylingDoc.FontFamilySize(NurtureDatarun3);
				NurtureData3.setNumID(stylingDoc.bullet(document));
				NurtureDatarun3.setText("Add new content and entire programs to nurture streams.");

				XWPFParagraph NurtureData4 = document.createParagraph();
				XWPFRun NurtureDatarun4 = NurtureData4.createRun();
				stylingDoc.FontFamilySize(NurtureDatarun4);
				NurtureData4.setNumID(stylingDoc.bullet(document));
				NurtureDatarun4.setText("Edit the availability of content.");

				XWPFParagraph NurtureData5 = document.createParagraph();
				XWPFRun NurtureDatarun5 = NurtureData5.createRun();
				stylingDoc.FontFamilySize(NurtureDatarun5);
				NurtureData5.setNumID(stylingDoc.bullet(document));
				NurtureDatarun5
						.setText("Understand content performance based on engagement with each piece of content.");

				XWPFParagraph Nurture_img = document.createParagraph();
				logger.info("Printing images in doc");
				XWPFRun img1 = Nurture_img.createRun();
				for (int i = 1; i < 2; i++) {
					try {

						img1.addPicture(new FileInputStream(passData.FetchScreenshot("Nurture campaigns")),
								Document.PICTURE_TYPE_PNG, passData.FetchScreenshot("Nurture campaigns").toString(),
								Units.toEMU(220), Units.toEMU(380));

					} catch (Exception e) {
						// TODO: handle exception
						logger.warn("Test is failed to retrive models images\n");
					}
				}

				XWPFParagraph NurtureLink = document.createParagraph();
				XWPFRun NurtureLinkdata = NurtureLink.createRun();
				stylingDoc.FontFamilySize(NurtureLinkdata);
				NurtureLinkdata.addCarriageReturn();
				NurtureLinkdata.setText("Please share the following resources with your customer: \r\n");
				NurtureLinkdata.addCarriageReturn();
				NurtureLinkdata.setText("\nMarketo Docs: Understanding Engagement Programs:");

				XWPFRun NurtureLink1 = NurtureLink.createRun();
				stylingDoc.FontFamilySize(NurtureLink1);
				stylingDoc.setNoProof(NurtureLink1);
				NurtureLink1.setUnderline(UnderlinePatterns.SINGLE);
				NurtureLink1.setColor("3333cc");
				NurtureLink1.setText(
						"\nhttps://experienceleague.adobe.com/docs/marketo/using/product-docs/email-marketing/drip-nurturing/creating-an-engagement-program/understanding-engagement-programs.html?lang=en");
				NurtureLink1.addCarriageReturn();

				XWPFRun NurtureLinkdata1 = NurtureLink.createRun();
				stylingDoc.FontFamilySize(NurtureLinkdata1);
				stylingDoc.setNoProof(NurtureLinkdata1);
				NurtureLinkdata1.setText("Marketo Docs: Create an Engagement Program: ");

				XWPFRun NurtureLink2 = NurtureLink.createRun();
				stylingDoc.FontFamilySize(NurtureLink2);
				stylingDoc.setNoProof(NurtureLink2);
				NurtureLink2.setUnderline(UnderlinePatterns.SINGLE);
				NurtureLink2.setColor("3333cc");
				NurtureLink2.setText(
						"\nhttps://experienceleague.adobe.com/docs/marketo/using/product-docs/email-marketing/drip-nurturing/creating-an-engagement-program/create-an-engagement-program.html?lang=en");

				XWPFRun NurtureLinkdata2 = NurtureLink.createRun();
				stylingDoc.FontFamilySize(NurtureLinkdata2);
				stylingDoc.setNoProof(NurtureLinkdata2);
				NurtureLinkdata2.addCarriageReturn();
				NurtureLinkdata2.setText("Marketo Best Practices – Pathway to Nurture ");

				XWPFRun NurtureLink3 = NurtureLink.createRun();
				stylingDoc.FontFamilySize(NurtureLink3);
				stylingDoc.setNoProof(NurtureLink3);
				NurtureLink3.setUnderline(UnderlinePatterns.SINGLE);
				NurtureLink3.setColor("3333cc");
				NurtureLink3.setText("\nhttps://go.marketo.com/nurture-resources.html");

				XWPFRun NurtureLinkdata3 = NurtureLink.createRun();
				stylingDoc.FontFamilySize(NurtureLinkdata3);
				stylingDoc.setNoProof(NurtureLinkdata3);
				NurtureLinkdata3.addCarriageReturn();
				NurtureLinkdata3.setText("Marketo University: Lead Nurturing: ");

				XWPFRun NurtureLink4 = NurtureLink.createRun();
				stylingDoc.FontFamilySize(NurtureLink4);
				stylingDoc.setNoProof(NurtureLink4);
				NurtureLink4.setUnderline(UnderlinePatterns.SINGLE);
				NurtureLink4.setColor("3333cc");
				NurtureLink4
						.setText("\nhttps://www.marketo.com/education/training/email-marketing/#lead-nurturing/learn");

				XWPFRun NurtureLinkdata4 = NurtureLink.createRun();
				stylingDoc.FontFamilySize(NurtureLinkdata4);
				stylingDoc.setNoProof(NurtureLinkdata4);
				NurtureLinkdata4.addCarriageReturn();
				NurtureLinkdata4.setText("Nurture Campaign Setup Discussion:");

				XWPFRun NurtureLink5 = NurtureLink.createRun();
				stylingDoc.FontFamilySize(NurtureLink5);
				stylingDoc.setNoProof(NurtureLink5);
				NurtureLink5.setUnderline(UnderlinePatterns.SINGLE);
				NurtureLink5.setColor("3333cc");
				NurtureLink5.setText(
						"\nhttps://nation.marketo.com/t5/product-discussions/nurture-campaign-setup/td-p/124770");

				XWPFRun NurtureLinkdata5 = NurtureLink.createRun();
				stylingDoc.FontFamilySize(NurtureLinkdata5);
				stylingDoc.setNoProof(NurtureLinkdata5);
				NurtureLinkdata5.addCarriageReturn();
				NurtureLinkdata5.setText("Creating an Adaptable, Scalable Nurture Program Template in Marketo:");

				XWPFRun NurtureLink6 = NurtureLink.createRun();
				stylingDoc.FontFamilySize(NurtureLink6);
				stylingDoc.setNoProof(NurtureLink6);
				NurtureLink6.setUnderline(UnderlinePatterns.SINGLE);
				NurtureLink6.setColor("3333cc");
				NurtureLink6.setText(
						"\n: https://nation.marketo.com/t5/champion-program-blogs/creating-an-adaptable-scalable-nurture-program-template-in/ba-p/300442");

			} else {
				NurtureDataRun = NurtureData.createRun();
				stylingDoc.FontFamilySize(NurtureDataRun);
				stylingDoc.setNoProof(NurtureDataRun);
				NurtureDataRun.addCarriageReturn();
				NurtureDataRun.setText(passData.Exceldata("Account Name") + passData.No_Nurture);

				XWPFParagraph NurtureData1 = document.createParagraph();
				XWPFRun NurtureDatarun1 = NurtureData1.createRun();
				stylingDoc.FontFamilySize(NurtureDatarun1);
				NurtureData1.setNumID(stylingDoc.bullet(document));
				NurtureData1.setSpacingAfter(0);
				NurtureDatarun1.setText("Intelligently and automatically deliver content to a target audience.");

				XWPFParagraph NurtureData2 = document.createParagraph();
				XWPFRun NurtureDatarun2 = NurtureData2.createRun();
				stylingDoc.FontFamilySize(NurtureDatarun2);
				NurtureData2.setSpacingAfter(0);
				NurtureData2.setNumID(stylingDoc.bullet(document));
				NurtureDatarun2.setText(
						"Easily build dialogue with prospects and customers while preventing customers who have already received content from receiving the same content again.");

				XWPFParagraph NurtureData3 = document.createParagraph();
				XWPFRun NurtureDatarun3 = NurtureData3.createRun();
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
				XWPFRun NurtureDatarun5 = NurtureData5.createRun();
				stylingDoc.FontFamilySize(NurtureDatarun5);
				NurtureData5.setNumID(stylingDoc.bullet(document));
				NurtureDatarun5
						.setText("Understand content performance based on engagement with each piece of content.");
				NurtureDatarun5.addCarriageReturn();

				XWPFParagraph NurtureLink = document.createParagraph();
				XWPFRun NurtureLinkdata = NurtureLink.createRun();
				stylingDoc.FontFamilySize(NurtureLinkdata);
				NurtureLinkdata.addCarriageReturn();
				NurtureLinkdata.setText("Please share the following resources with your customer: \r\n");
				NurtureLinkdata.addCarriageReturn();
				NurtureLinkdata.setText("\nMarketo Docs: Understanding Engagement Programs:");

				XWPFRun NurtureLink1 = NurtureLink.createRun();
				stylingDoc.FontFamilySize(NurtureLink1);
				stylingDoc.setNoProof(NurtureLink1);
				NurtureLink1.setUnderline(UnderlinePatterns.SINGLE);
				NurtureLink1.setColor("3333cc");
				NurtureLink1.setText(
						"\nhttps://experienceleague.adobe.com/docs/marketo/using/product-docs/email-marketing/drip-nurturing/creating-an-engagement-program/understanding-engagement-programs.html?lang=en");
				NurtureLink1.addCarriageReturn();

				XWPFRun NurtureLinkdata1 = NurtureLink.createRun();
				stylingDoc.FontFamilySize(NurtureLinkdata1);
				stylingDoc.setNoProof(NurtureLinkdata1);
				NurtureLinkdata1.setText("Marketo Docs: Create an Engagement Program: ");

				XWPFRun NurtureLink2 = NurtureLink.createRun();
				stylingDoc.FontFamilySize(NurtureLink2);
				stylingDoc.setNoProof(NurtureLink2);
				NurtureLink2.setUnderline(UnderlinePatterns.SINGLE);
				NurtureLink2.setColor("3333cc");
				NurtureLink2.setText(
						"\nhttps://experienceleague.adobe.com/docs/marketo/using/product-docs/email-marketing/drip-nurturing/creating-an-engagement-program/create-an-engagement-program.html?lang=en");

				XWPFRun NurtureLinkdata2 = NurtureLink.createRun();
				stylingDoc.FontFamilySize(NurtureLinkdata2);
				stylingDoc.setNoProof(NurtureLinkdata2);
				NurtureLinkdata2.addCarriageReturn();
				NurtureLinkdata2.setText("Marketo Best Practices – Pathway to Nurture ");

				XWPFRun NurtureLink3 = NurtureLink.createRun();
				stylingDoc.FontFamilySize(NurtureLink3);
				stylingDoc.setNoProof(NurtureLink3);
				NurtureLink3.setUnderline(UnderlinePatterns.SINGLE);
				NurtureLink3.setColor("3333cc");
				NurtureLink3.setText("\nhttps://go.marketo.com/nurture-resources.html");

				XWPFRun NurtureLinkdata3 = NurtureLink.createRun();
				stylingDoc.FontFamilySize(NurtureLinkdata3);
				stylingDoc.setNoProof(NurtureLinkdata3);
				NurtureLinkdata3.addCarriageReturn();
				NurtureLinkdata3.setText("Marketo University: Lead Nurturing: ");

				XWPFRun NurtureLink4 = NurtureLink.createRun();
				stylingDoc.FontFamilySize(NurtureLink4);
				stylingDoc.setNoProof(NurtureLink4);
				NurtureLink4.setUnderline(UnderlinePatterns.SINGLE);
				NurtureLink4.setColor("3333cc");
				NurtureLink4
						.setText("\nhttps://www.marketo.com/education/training/email-marketing/#lead-nurturing/learn");

				XWPFRun NurtureLinkdata4 = NurtureLink.createRun();
				stylingDoc.FontFamilySize(NurtureLinkdata4);
				stylingDoc.setNoProof(NurtureLinkdata4);
				NurtureLinkdata4.addCarriageReturn();
				NurtureLinkdata4.setText("Nurture Campaign Setup Discussion:");

				XWPFRun NurtureLink5 = NurtureLink.createRun();
				stylingDoc.FontFamilySize(NurtureLink5);
				stylingDoc.setNoProof(NurtureLink5);
				NurtureLink5.setUnderline(UnderlinePatterns.SINGLE);
				NurtureLink5.setColor("3333cc");
				NurtureLink5.setText(
						"\nhttps://nation.marketo.com/t5/product-discussions/nurture-campaign-setup/td-p/124770");

				XWPFRun NurtureLinkdata5 = NurtureLink.createRun();
				stylingDoc.FontFamilySize(NurtureLinkdata5);
				stylingDoc.setNoProof(NurtureLinkdata5);
				NurtureLinkdata5.addCarriageReturn();
				NurtureLinkdata5.setText("Creating an Adaptable, Scalable Nurture Program Template in Marketo:");

				XWPFRun NurtureLink6 = NurtureLink.createRun();
				stylingDoc.FontFamilySize(NurtureLink6);
				stylingDoc.setNoProof(NurtureLink6);
				NurtureLink6.setUnderline(UnderlinePatterns.SINGLE);
				NurtureLink6.setColor("3333cc");
				NurtureLink6.setText(
						"\n: https://nation.marketo.com/t5/champion-program-blogs/creating-an-adaptable-scalable-nurture-program-template-in/ba-p/300442");

			}

			logger.info(" Nurture data part is done......");
		} catch (NumberFormatException ex) {
			logger.error("Nurture Data Test is failed Hence we have null value\n");
			// handle your exception

		}

	}

	public static void segment(XWPFDocument document)
			throws NumberFormatException, IOException, InvalidFormatException {

		XWPFParagraph Segment = document.createParagraph();
		XWPFRun SegmentHeading = Segment.createRun();
		SegmentHeading.setBold(true);
		logger.info("Printing Segmentation Data");
		SegmentHeading.setText("Segmentation");
		try {
			int Segment_Data = Integer.parseInt(passData.Exceldata("Segmentations"));
			if (Segment_Data > 0) {

				XWPFParagraph SegmentData = document.createParagraph();
				XWPFRun SegmentRun = SegmentData.createRun();
				stylingDoc.FontFamilySize(SegmentRun);
				stylingDoc.setNoProof(SegmentRun);
				SegmentRun.setText(passData.Exceldata("Account Name") + passData.segment);

				SegmentRun = SegmentData.createRun();

				SegmentRun.addBreak();

				for (int i = 1; i < 4; i++) {
					try {
						SegmentRun.addCarriageReturn();
						SegmentRun.addPicture(new FileInputStream(passData.FetchScreenshot("Segmentations" + i)),
								Document.PICTURE_TYPE_PNG, passData.FetchScreenshot("Segmentations"), Units.toEMU(500),
								Units.toEMU(80));
						SegmentRun.addCarriageReturn();
					} catch (Exception e) {
						// TODO: handle exception
						logger.warn("We dont find any images so it does not printing images\n");
					}
				}
			} else {

				XWPFParagraph InterestingData = document.createParagraph();
				XWPFRun InterestingDatarun = InterestingData.createRun();
				stylingDoc.FontFamilySize(InterestingDatarun);
				stylingDoc.setNoProof(InterestingDatarun);
				InterestingDatarun.setText(passData.no_segment);

				XWPFParagraph segment = document.createParagraph();
				XWPFRun segmentalalternet = segment.createRun();
				stylingDoc.FontFamilySize(segmentalalternet);
				segmentalalternet.addCarriageReturn();
				segmentalalternet.setText("I would direct them to the docs below:\r\n");
				segmentalalternet.addCarriageReturn();
				segmentalalternet.setText("\nMarketo Docs: Create a Segmentation: ");

				XWPFRun segmentLink = segment.createRun();
				stylingDoc.FontFamilySize(segmentLink);
				stylingDoc.setNoProof(segmentLink);
				segmentLink.setUnderline(UnderlinePatterns.SINGLE);
				segmentLink.setColor("3333cc");
				segmentLink.setText(
						"\nhttps://experienceleague.adobe.com/docs/marketo/using/product-docs/personalization/segmentation-and-snippets/segmentation/create-a-segmentation.html?lang=en");
				segmentLink.addCarriageReturn();

				XWPFRun segment2 = segment.createRun();
				stylingDoc.FontFamilySize(segment2);
				stylingDoc.setNoProof(segment2);
				segment2.setText(
						"Segmentation Health Check Updates – Tips and Tricks for Keeping Your Segmentation Updated: ");

				XWPFRun segmentLink2 = segment.createRun();
				stylingDoc.FontFamilySize(segmentLink2);
				stylingDoc.setNoProof(segmentLink2);
				segmentLink2.setUnderline(UnderlinePatterns.SINGLE);
				segmentLink2.setColor("3333cc");
				segmentLink2.setText(
						"\nhttps://nation.marketo.com/t5/product-blogs/segmentation-health-check-updates-tips-and-tricks-for-keeping/ba-p/241963");
				segmentLink2.addCarriageReturn();

				XWPFRun segment3 = segment.createRun();
				stylingDoc.FontFamilySize(segment3);
				stylingDoc.setNoProof(segment3);
				segment3.addCarriageReturn();
				segment3.setText("Marketo Segmentation Success Guide: ");

				XWPFRun segmentLink3 = segment.createRun();
				stylingDoc.FontFamilySize(segmentLink3);
				stylingDoc.setNoProof(segmentLink3);
				segmentLink3.setUnderline(UnderlinePatterns.SINGLE);
				segmentLink3.setColor("3333cc");
				segmentLink3.setText(
						"\nhttps://nation.marketo.com/t5/product-blogs/marketo-engage-segmentation-marketo-engage-success-guide-sneak/ba-p/242620");
				segmentLink3.addCarriageReturn();

				XWPFRun segment4 = segment.createRun();
				stylingDoc.FontFamilySize(segment4);
				stylingDoc.setNoProof(segment4);
				segment4.addCarriageReturn();
				segment4.setText("Marketo Docs: Define Segment Rules: ");

				XWPFRun segmentLink4 = segment.createRun();
				stylingDoc.FontFamilySize(segmentLink4);
				stylingDoc.setNoProof(segmentLink4);
				segmentLink4.setUnderline(UnderlinePatterns.SINGLE);
				segmentLink4.setColor("3333cc");
				segmentLink4.setText(
						"\nhttps://experienceleague.adobe.com/docs/marketo/using/product-docs/personalization/segmentation-and-snippets/segmentation/define-segment-rules.html");

				XWPFRun segment5 = segment.createRun();
				stylingDoc.FontFamilySize(segment5);
				stylingDoc.setNoProof(segment5);
				segment5.addCarriageReturn();
				segment5.setText("Marketo Docs: Understanding Dynamic Content: ");

				XWPFRun segmentLink5 = segment.createRun();
				stylingDoc.FontFamilySize(segmentLink5);
				stylingDoc.setNoProof(segmentLink5);
				segmentLink5.setUnderline(UnderlinePatterns.SINGLE);
				segmentLink5.setColor("3333cc");
				segmentLink5.setText(
						"\nhttps://experienceleague.adobe.com/docs/marketo/using/product-docs/personalization/segmentation-and-snippets/segmentation/understanding-dynamic-content.html");

				XWPFRun segment6 = segment.createRun();
				stylingDoc.FontFamilySize(segment6);
				stylingDoc.setNoProof(segment6);
				segment6.addCarriageReturn();
				segment6.setText("Segment or Filter?: ");

				XWPFRun segmentLink6 = segment.createRun();
				stylingDoc.FontFamilySize(segmentLink6);
				stylingDoc.setNoProof(segmentLink4);
				segmentLink6.setUnderline(UnderlinePatterns.SINGLE);
				segmentLink6.setColor("3333cc");
				segmentLink6.setText(
						"\nhttps://nation.marketo.com/t5/product-discussions/segment-or-filter/td-p/149738\r\n");
				segmentLink6.addCarriageReturn();

				XWPFRun segment7 = segment.createRun();
				stylingDoc.FontFamilySize(segment7);
				stylingDoc.setNoProof(segment7);
				segment7.addCarriageReturn();
				segment7.setText("How to Effectively Segment Your Database for Lead Nurturing: ");

				XWPFRun segmentLink7 = segment.createRun();
				stylingDoc.FontFamilySize(segmentLink7);
				stylingDoc.setNoProof(segmentLink7);
				segmentLink7.setUnderline(UnderlinePatterns.SINGLE);
				segmentLink7.setColor("3333cc");
				segmentLink7.setText(
						"\nhttps://blog.marketo.com/2013/10/how-to-effectively-segment-your-database-for-lead-nurturing.html");

				XWPFRun segment8 = segment.createRun();
				stylingDoc.FontFamilySize(segment8);
				stylingDoc.setNoProof(segment8);
				segment8.addCarriageReturn();
				segment8.setText("Marketo eBook: How to Segment and Target Your Emails: ");

				XWPFRun segmentLink8 = segment.createRun();
				stylingDoc.FontFamilySize(segmentLink8);
				stylingDoc.setNoProof(segmentLink8);
				segmentLink8.setUnderline(UnderlinePatterns.SINGLE);
				segmentLink8.setColor("3333cc");
				segmentLink8.setText("\nhttps://www.marketo.com/ebooks/how-to-segment-and-target-your-emails/");

			}
		} catch (NumberFormatException ex) { // handle your exception
			logger.error("Test is skipped or failed so we have null value here\n");
		}
		logger.info(" Segmentation data part is done......");

	}

	public static void programLibrary(XWPFDocument document) throws IOException, InvalidFormatException {
		try {
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
			for (int i = 1; i < 5; i++) {
				try {
					ProgramDataRun.addPicture(new FileInputStream(passData.FetchScreenshot("Library" + i)),
							Document.PICTURE_TYPE_PNG, passData.FetchScreenshot("Library" + i).toString(),
							Units.toEMU(470), Units.toEMU(200));
				}

				catch (Exception e) {
					// TODO: handle exception
					logger.warn("There is no image in Program Library");
				}
			}
		} catch (Exception e) {
			// TODO: handle exception
		}
	}

	public static void Integration(XWPFDocument document) throws IOException, InvalidFormatException {
		XWPFParagraph intigration = document.createParagraph();
		XWPFRun intigrationRunHeading = intigration.createRun();
		intigrationRunHeading.addCarriageReturn();
		intigrationRunHeading.addCarriageReturn();
		intigrationRunHeading.setBold(true);
		logger.info("Printing Integration Data");
		intigrationRunHeading.setText("Integrations");
		try {
			int intigration_Data = Integer.parseInt(passData.Exceldata("Integration"));
			if (intigration_Data > 0) {

				XWPFParagraph intigrationRun = document.createParagraph();
				XWPFRun intigrationData = intigrationRun.createRun();
				stylingDoc.FontFamilySize(intigrationData);
				intigrationData.setText("The following integrations have been installed:");

				intigrationRun = document.createParagraph();
				XWPFRun intigrationImg = intigrationRun.createRun();
				for (int i = 1; i < 2; i++) {
					try {
						intigrationImg.addCarriageReturn();
						intigrationImg.addPicture(new FileInputStream(passData.FetchScreenshot("Integration")),
								Document.PICTURE_TYPE_PNG, passData.FetchScreenshot("Integration"), Units.toEMU(470),
								Units.toEMU(200));
						intigrationImg.addCarriageReturn();
					} catch (Exception e) {
						logger.warn("we dont have images to print\n");
						// TODO: handle exception
					}
				}

			} else {

				XWPFParagraph intigrationRun = document.createParagraph();
				XWPFRun intigrationData = intigrationRun.createRun();
				stylingDoc.FontFamilySize(intigrationData);
				stylingDoc.setNoProof(intigrationData);

				intigrationData.setText(passData.Exceldata("Account Name") + passData.No_Integrations);

				XWPFRun intigrationData1 = intigrationRun.createRun();
				stylingDoc.FontFamilySize(intigrationData1);
				stylingDoc.setNoProof(intigrationData1);
				intigrationData1.addCarriageReturn();
				intigrationData1.addCarriageReturn();
				intigrationData1.setText(
						"Here is a starting point for the client on some common Marketo Partners and step-by-step instructions on installing each one: ");

				XWPFRun intigrationDatalink = intigrationRun.createRun();
				stylingDoc.FontFamilySize(intigrationDatalink);
				stylingDoc.setNoProof(intigrationDatalink);
				intigrationDatalink.addCarriageReturn();
				intigrationDatalink.setUnderline(UnderlinePatterns.SINGLE);
				intigrationDatalink.setColor("3333cc");
				intigrationDatalink
						.setText("\nhttps://docs.marketo.com/display/public/DOCS/LaunchPoint+Event+Partners.");
			}

		} catch (NumberFormatException ex) { // handle your exception
			logger.error("Intigration Test is skiped\n");
		}

	}

	public static void webPersonalize(XWPFDocument document) throws IOException, InvalidFormatException {
		try {
			if (passData.Exceldata("Web Personalize").equalsIgnoreCase("true")) {
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
				PersonalizeImg.setText(
						"Top Campaigns: The top performing campaigns during the selected time period, ordered by number of clicks.");

				XWPFRun PersonalizeImg1 = PersonalizeDoc.createRun();
				stylingDoc.FontFamilySize(PersonalizeImg1);
				stylingDoc.setNoProof(PersonalizeImg1);
				PersonalizeImg1.addCarriageReturn();
				PersonalizeImg1.addPicture(new FileInputStream(passData.FetchScreenshot("Change Score1")),
						Document.PICTURE_TYPE_PNG, passData.FetchScreenshot("Change Score"), Units.toEMU(470),
						Units.toEMU(80));
				PersonalizeImg1.addBreak();
				PersonalizeImg1.addCarriageReturn();
				PersonalizeImg1.addPicture(new FileInputStream(passData.FetchScreenshot("Change Score1")),
						Document.PICTURE_TYPE_PNG, passData.FetchScreenshot("Change Score"), Units.toEMU(470),
						Units.toEMU(80));
				PersonalizeImg1.addBreak();
				PersonalizeImg1.addCarriageReturn();
				PersonalizeImg1.addPicture(new FileInputStream(passData.FetchScreenshot("Change Score1")),
						Document.PICTURE_TYPE_PNG, passData.FetchScreenshot("Change Score"), Units.toEMU(470),
						Units.toEMU(80));
				PersonalizeImg1.addCarriageReturn();
				PersonalizeImg1.setText(
						"Total Organizations & Top 5 Organizations: The Organizations tab displays all the details (name, location, activity and time stamp) of organizations that visited your"
								+ " website during a given period. The table can be sorted and organized by time, location, domain and via a free text search.");

				// PersonalizeImg1.addCarriageReturn();
				PersonalizeImg1.addPicture(new FileInputStream(passData.FetchScreenshot("Change Score1")),
						Document.PICTURE_TYPE_PNG, passData.FetchScreenshot("Change Score"), Units.toEMU(470),
						Units.toEMU(80));
				PersonalizeImg1.addCarriageReturn();
			}
		} catch (Exception e) {
			// TODO: handle exception
		}
	}

	public static void TAM(XWPFDocument document) throws IOException, InvalidFormatException {
		try {
			if (passData.Exceldata("Target Account Management").equalsIgnoreCase("true")) {
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
					OverviewRun.addPicture(new FileInputStream(passData.FetchScreenshot("Integration")),
							Document.PICTURE_TYPE_PNG, passData.FetchScreenshot("Integration").toString(),
							Units.toEMU(470), Units.toEMU(80));
					OverviewRun.addBreak();
					OverviewRun.setText("Top Named Accounts (by Pipeline)");
					OverviewRun.addBreak();
					OverviewRun.addPicture(new FileInputStream(passData.FetchScreenshot("Integration")),
							Document.PICTURE_TYPE_PNG, passData.FetchScreenshot("Integration").toString(),
							Units.toEMU(470), Units.toEMU(280));
				} catch (Exception e) {
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
				HigestAccountRunData.setText(
						"Marketo’s ABM Account Score is a systematic approach designed to help Sales and Marketing teams identify and prioritize "
								+ "the companies (including prospects) that are most likely to make a purchase.");
				HigestAccountRunData.addCarriageReturn();
				HigestAccountRunData.addCarriageReturn();

				HigestAccountRunData.setText(
						"In the complex world of B2B buying processes, it’s rare that a single individual makes a purchase decision. "
								+ "There are often various roles involved, each with their own needs. Account-based scoring takes this into account by aggregating the"
								+ " lead scores from multiple leads and providing a score at an account level.\n"
								+ passData.Exceldata("Account Name") + "’s highest account score is from Cisco with a\n"
								+ "989.");
				XWPFRun HigestAccountRunimg = HigestAccount.createRun();
				HigestAccountRunimg.addPicture(new FileInputStream(passData.FetchScreenshot("Integration")),
						Document.PICTURE_TYPE_PNG, passData.FetchScreenshot("Integration").toString(), Units.toEMU(470),
						Units.toEMU(110));
				HigestAccountRunimg.setBold(true);
				HigestAccountRunimg.setText("Account Lists");
				try {
					HigestAccountRunimg.addCarriageReturn();
					HigestAccountRunimg.addPicture(new FileInputStream(passData.FetchScreenshot("Integration")),
							Document.PICTURE_TYPE_PNG, passData.FetchScreenshot("Integration").toString(),
							Units.toEMU(470), Units.toEMU(110));
				} catch (Exception e) {
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
				LargestAccountListRunData.setText(
						"An account list is a collection of named accounts that can be targeted together. Account lists allow you to target named accounts by industry, location or size of the company. Marketo allows for up to 2,000 named accounts to be added to an account list.\n"
								+ passData.Exceldata("Account Name")
								+ "\n’s largest account list is NA Target Accounts, containing 545 named accounts and is responsible for $23.1 million in pipeline.");

				XWPFRun LargestAccountListRunimg = LargestAccountList.createRun();
				try {
					LargestAccountListRunimg.addPicture(new FileInputStream(passData.FetchScreenshot("Integration")),
							Document.PICTURE_TYPE_PNG, passData.FetchScreenshot("Integration").toString(),
							Units.toEMU(470), Units.toEMU(80));
				} catch (Exception e) {
					logger.warn("Test is skipped so user dont have images to print\n");
					// TODO: handle exception
				}
				LargestAccountListRunimg.addCarriageReturn();
				LargestAccountListRunimg.addCarriageReturn();
				stylingDoc.FontFamilySize(LargestAccountListRunimg);
				stylingDoc.setNoProof(LargestAccountListRunimg);
				LargestAccountListRunimg.setText(
						"To learn more about using Account Lists in Marketo, visit: http://docs.marketo.com/display/public/DOCS/Account+Lists");

			}
		} catch (Exception e) {
			// TODO: handle exception
		}

	}

	public static void PredictiveContent(XWPFDocument document) throws IOException, InvalidFormatException {
		try {
			if (passData.Exceldata("Predictive Content").equalsIgnoreCase("true")) {
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

		} catch (Exception e) {
			// TODO: handle exception
		}
	}

}
