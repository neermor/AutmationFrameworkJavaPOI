package com.marketo.qa.utility;

import java.io.FileOutputStream;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Locale;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelARReport extends converting {

	static converting obj = new converting();

	static String ExcelPath = System.getProperty("user.dir") + "//Reports//";
	static String fileName = new SimpleDateFormat("dd_MM_yy_HH_mm").format(new Date());
	private static final Logger logger = LogManager.getLogger(ExcelARReport.class);
	public static void ARReport() throws Exception {
		
		converting.localtest();

		String AR = System.getProperty("user.dir") + "//Config//AR_Temp.xlsx";
		String report = System.getProperty("user.dir") + "//Reports//" + passData.Exceldata("Account Name") + "_"
				+ fileName + ".xlsx";
		// FileInputStream FileInputStream = new FileInputStream(report);
		logger.info("Creating blank templet");
		// Create a blank sheet
		XSSFWorkbook workbookinput = new XSSFWorkbook(AR);
		XSSFWorkbook workbookoutput = workbookinput;

		int WorkSpace_count = Integer.parseInt(passData.Exceldata("Total WorkSpace"));

		XSSFSheet AR_Data_Point1 = workbookoutput.getSheet("AR Data Points 1");
		XSSFSheet AR_Data_Point2 = workbookoutput.getSheet("AR Data Points 2");

		// Adding prefix name
		logger.info("Adding Prifix in account");
		ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point1, 1, 1, "FFFFFF", "Requested By CSM", 11);
		// adjusting column width for workspace

		AR_Data_Point1.setColumnWidth(3 + WorkSpace_count, 25 * 256);
		AR_Data_Point1.setColumnWidth(5 + WorkSpace_count, 25 * 1000);
		logger.info("Seting column width");
		// inserting tag Images
		int row = 47;
		for (int i = 1; i <= 1; i++) {

			try {
				logger.info("Printing Tag screenshot");
				ExcelStyling.pestImg(workbookoutput, AR_Data_Point1, 2, row, passData.FetchScreenshot("Tags" + i),0.6);
				row = row + 22;

			} catch (Exception e) {
				logger.info("Image not available");
				// TODO: handle exception
			}

		}

		
		// printing data in sheet
		if (WorkSpace_count >= 1) {
			logger.info("Seting templete according to Workspace count");
				AR_Data_Point1.shiftColumns(3, 9, WorkSpace_count);

				// Printing the Heading
				try {
				ExcelStyling.mergeAndCenter(workbookoutput, AR_Data_Point1, "D662FC", "AR Data Points", 1, 2, 2,
						3 + WorkSpace_count, true, 25);

				ExcelStyling.mergeAndCenter(workbookoutput, AR_Data_Point1, "B541DF",
						"instance acct prefix: " + passData.Exceldata("Account Name"), 3, 3, 2, 3 + WorkSpace_count,
						false, 12);
				}
				catch (Exception e) {
					// TODO: handle exception
				logger.warn("not able to merge and print account prifix");
				}
				int coll = 3;
				int j = 2;
				int k = 7;

				// Printing workspaces names
				for (int i = 1; i <= WorkSpace_count; i++) {
					logger.info("Printing workspaces");
//						ExcelStyling.mergeAndCenter(workbookoutput, AR_Data_Point1, "D662FC", "AR Data Points", 1, 2, 2,
//								3 + WorkSpace_count, true, 25);
//
//						ExcelStyling.mergeAndCenter(workbookoutput, AR_Data_Point1, "B541DF",
//								"instance acct prefix: " + passData.Exceldata("Account Name"), 3, 3, 2,
//								3 + WorkSpace_count, false, 12);

						AR_Data_Point1.setColumnWidth(coll, 25 * 256);
						ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point1, 6, coll++, "D662FC",
								passData.Exceldata("Asset Data", j++), 14);
						j--;
						coll--;
						ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point1, 17, coll++, "D662FC",
								passData.Exceldata("Asset Data", j++), 14);
						j--;
						coll--;

						ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point1, 26, coll++, "D662FC",
								passData.Exceldata("Asset Data", j++), 14);

						ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point1, 40, 3, "D662FC",
								passData.Exceldata("Asset Data"), 12);

					
				}

				// Campaign Data
				j = 2;
				coll = 3;
				// k =7;
				for (int i = 1; i <= WorkSpace_count; i++) {
					logger.info("Printing Campaign data in coll"+coll);
					try {
					//	int AllCampaigns = Integer.parseInt(passData.Exceldata("All Campaigns", j++));
						ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point1, 7, coll++, passData.Exceldata("All Campaigns", j++), 12);
						coll--;
						j--;
						
					}  
					   catch (Exception e) {
						// TODO: handle exception
					}

					try {
					//	int ActiveCampaigns = Integer.parseInt(passData.Exceldata("Active Campaigns", j++));
						ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point1, 8, coll++, passData.Exceldata("Active Campaigns", j++), 12);
						coll--;
						j--;
						
					}  
					catch (Exception e) {
						// TODO: handle exception
					}
					try {
					//	int AllTriggeredCampaigns = Integer.parseInt(passData.Exceldata("All Triggered Campaigns", j++));
						ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point1, 9, coll++, passData.Exceldata("All Triggered Campaigns", j++), 12);
						coll--;
						j--;
					} catch (Exception e) {
						// TODO: handle exception
					}

					try {
						//int ActiveTriggeredCampaigns = Integer.parseInt(passData.Exceldata("Active Triggered Campaigns", j++));
						ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point1, 10, coll++, passData.Exceldata("Active Triggered Campaigns", j++),
								12);
						coll--;
						j--;
					} catch (Exception e) {
						// TODO: handle exception
					}
					try {

						//int AllBatchCampaigns = Integer.parseInt(passData.Exceldata("All Batch Campaigns", j++));
						ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point1, 11, coll++, passData.Exceldata("All Batch Campaigns", j++), 12);
						coll--;
						j--;
					} catch (Exception e) {
						// TODO: handle exception
					}
					try {
					//	int BatchCampaigns = Integer.parseInt(passData.Exceldata("Batch Campaigns - Repeating Schedule", j++));
						ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point1, 12, coll++, passData.Exceldata("Batch Campaigns - Repeating Schedule", j++), 12);
						coll--;
						j--;
					} catch (Exception e) {
						// TODO: handle exception
					}
					try {
					//	int ScheduledBatchCampaigns = Integer.parseInt(passData.Exceldata("Scheduled Batch Campaigns", j++));
						ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point1, 13, coll++, passData.Exceldata("Scheduled Batch Campaigns", j++),
								12);
						coll--;
						j--;
					} catch (Exception e) {
						// TODO: handle exception
					}
				}

				// Asset Data printing
				try {
					//int LandingPages = Integer.parseInt(passData.Exceldata("Landing Pages", j++));
					ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point1, 18, coll++, passData.Exceldata("Landing Pages", j++), 12);
					coll--;
					j--;
				} catch (Exception e) {
					// TODO: handle exception
				}

				try {
					//int Forms = Integer.parseInt(passData.Exceldata("Forms", j++));
					ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point1, 19, coll++, passData.Exceldata("Forms", j++), 12);
					coll--;
					j--;
				} catch (Exception e) {
					// TODO: handle exception
				}

				try {
					//int Emails = Integer.parseInt(passData.Exceldata("Emails", j++));
			
					ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point1, 20, coll++, passData.Exceldata("Emails", j++), 12);
					coll--;
					j--;
				} catch (Exception e) {
					// TODO: handle exception
				}

				try {
					//int Snippets = Integer.parseInt(passData.Exceldata("Snippets", j++));
					
					ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point1, 21, coll++, passData.Exceldata("Snippets", j++), 12);
					coll--;
					j--;
				} catch (Exception e) {
					// TODO: handle exception
				}
				try {
				//	int ImagesandFiles = Integer.parseInt(passData.Exceldata("Images and Files", j++));
					ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point1, 22, coll++, passData.Exceldata("Images and Files", j++), 12);
					coll--;
					j--;
				} catch (Exception e) {
					// TODO: handle exception
				}
				// Database Data
				try {
				//	int AllPeople = Integer.parseInt(passData.Exceldata("All People", j++));
					ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point1, 27, coll++, passData.Exceldata("All People", j++), 12); // Need to check
					coll--;
					j--;
				} catch (Exception e) {
					// TODO: handle exception
				}
				try {
					
					ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point1, 28, coll++, passData.Exceldata("Marketable Leads", j++), 12); // Need to
																											
					coll--;
					j--;
				} catch (Exception e) {
					// TODO: handle exception
				}
				try {
					//int PossibleDuplicates = Integer.parseInt(passData.Exceldata("Possible Duplicates", j++));
					ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point1, 29, coll++, passData.Exceldata("Possible Duplicates", j++), 12);
					coll--;
					j--;
				} catch (Exception e) {
					// TODO: handle exception
				}
				try {
					//int UnsubscribedPeople = Integer.parseInt(passData.Exceldata("Unsubscribed People", j++));
					ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point1, 30, coll++, passData.Exceldata("Unsubscribed People", j++), 12);
					coll--;
					j--;
				} catch (Exception e) {
					// TODO: handle exception
				}
				try {
				//	int MarketingSuspended = Integer.parseInt(passData.Exceldata("Marketing Suspended", j++));
					ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point1, 31, coll++, passData.Exceldata("Marketing Suspended", j++), 12);
				} catch (Exception e) {
					// TODO: handle exception
				}
				
				

				// Totals of Campaigns
				try {
					ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point1, 7, 3 + WorkSpace_count, passData.Exceldata("All Campaigns"), 12);
				} catch (Exception e) {
					// TODO: handle exception
				}
				try {
					ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point1, 8, 3 + WorkSpace_count, passData.Exceldata("Active Campaigns"),
							12);
				} catch (Exception e) {
					// TODO: handle exception
				}
				
				try {
					ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point1, 9, 3 + WorkSpace_count,
							passData.Exceldata("All Triggered Campaigns"), 12);
				} catch (Exception e) {
					// TODO: handle exception
				}
				try {
					ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point1, 10, 3 + WorkSpace_count,
							passData.Exceldata("Active Triggered Campaigns"), 12);
				} catch (Exception e) {
					// TODO: handle exception
				}

				try {
					ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point1, 11, 3 + WorkSpace_count, passData.Exceldata("All Batch Campaigns"),
							12);
				} catch (Exception e) {
					// TODO: handle exception
				}
				try {
					ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point1, 12, 3 + WorkSpace_count,
							passData.Exceldata("Batch Campaigns - Repeating Schedule"), 12);
				} catch (Exception e) {
					// TODO: handle exception
				}
				try {
					ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point1, 13, 3 + WorkSpace_count, passData.Exceldata("Scheduled Batch Campaigns"),
							12);
				} catch (Exception e) {
					// TODO: handle exception
				}

				// printing Asset Data totals for multiple workspaces
				try {
					ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point1, 18, 3 + WorkSpace_count, passData.Exceldata("Landing Pages"), 12);
				} catch (Exception e) {
					// TODO: handle exception
				}
				try {
					ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point1, 19, 3 + WorkSpace_count, passData.Exceldata("Forms"), 12);
				} catch (Exception e) {
					// TODO: handle exception
				}
				try {
					ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point1, 20, 3 + WorkSpace_count, passData.Exceldata("Emails"), 12);
				} catch (Exception e) {
					// TODO: handle exception
				}
				try {
					ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point1, 21, 3 + WorkSpace_count, passData.Exceldata("Snippets"), 12);
				} catch (Exception e) {
					// TODO: handle exception
				}
				try {
					ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point1, 22, 3 + WorkSpace_count, passData.Exceldata("Images and Files"),
							12);
				} catch (Exception e) {
					// TODO: handle exception
				}
				// printing Database Data totals for multiple workspaces
				try {
					ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point1, 27, 3 + WorkSpace_count, passData.Exceldata("All People"), 12); // Need
																														// to
																														// check
				} catch (Exception e) {
					// TODO: handle exception
				}
				try {
					ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point1, 28, 3 + WorkSpace_count, passData.Exceldata("Marketable Leads"),
							12); // Need to check
				} catch (Exception e) {
					// TODO: handle exception
				}
				try {
					ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point1, 29, 3 + WorkSpace_count, passData.Exceldata("Possible Duplicates"),
							12);
				} catch (Exception e) {
					// TODO: handle exception
				}
				try {

					ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point1, 30, 3 + WorkSpace_count, passData.Exceldata("Unsubscribed People"),
							12);
				} catch (Exception e) {
					// TODO: handle exception
				}
				try {
					ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point1, 31, 3 + WorkSpace_count, passData.Exceldata("Marketing Suspended"),
							12);
				} catch (Exception e) {
					// TODO: handle exception
				}
				
				
				try {
					float Marketable_leads =converting.MarketableLeads/(float) converting.AllPeople*100;
					System.out.println(Marketable_leads);
					//workbookoutput.createDataFormat().getFormat("00.00%");
					ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point1, 34, 3 + WorkSpace_count, new DecimalFormat("##.##").format(Marketable_leads)+"%",
							12);
				} catch (Exception e) {
					// TODO: handle exception
				}
				
				try {
					float Marketable_leads =converting.PossibleDuplicates/(float) converting.AllPeople*100;
					System.out.println(Marketable_leads);
					//workbookoutput.createDataFormat().getFormat("00.00%");
					ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point1, 35, 3 + WorkSpace_count, new DecimalFormat("##.##").format(Marketable_leads)+"%",
							12);
				} catch (Exception e) {
					// TODO: handle exception
				}
				
				try {
					float Marketable_leads =converting.UnsubscribedPeople/(float) converting.AllPeople*100;
					System.out.println(Marketable_leads);
					//workbookoutput.createDataFormat().getFormat("00.00%");
					ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point1, 36, 3 + WorkSpace_count, new DecimalFormat("##.##").format(Marketable_leads)+"%",
							12);
				} catch (Exception e) {
					// TODO: handle exception
				}
				
				try {
					float Marketable_leads =converting.MarketingSuspended/(float) converting.AllPeople*100;
					System.out.println(Marketable_leads);
					//workbookoutput.createDataFormat().getFormat("00.00%");
					ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point1, 37, 3 + WorkSpace_count, new DecimalFormat("##.##").format(Marketable_leads)+"%",
							12);
				} catch (Exception e) {
					// TODO: handle exception
				}
				
				// Printing Program Data
				try {
					ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point1, 41, 3, passData.Exceldata("Tags"), 12);
				} catch (Exception e) {
					// TODO: handle exception
				}
				try {
					ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point1, 42, 3, passData.Exceldata("Channels"), 12);
				} catch (Exception e) {
					// TODO: handle exception
				}
			
		}

		else {

			AR_Data_Point1.setColumnWidth(3, 25 * 256);
			ExcelStyling.mergeAndCenter(workbookoutput, AR_Data_Point1, "D662FC", "AR Data Points", 1, 2, 2,
					2 + WorkSpace_count, true, 25);

			ExcelStyling.mergeAndCenter(workbookoutput, AR_Data_Point1, "B541DF",
					"instance acct prefix: " + passData.Exceldata("Account Name"), 3, 3, 2, 2 + WorkSpace_count, false,
					12);

			ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point1, 7, 3, AllCampaigns, 12);

			ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point1, 8, 3, ActiveChampaigns, 12);

			ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point1, 9, 3, AllTriggeredCampaigns, 12);

			ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point1, 10, 3, ActiveTriggeredCampaigns, 12);

			ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point1, 11, 3, AllBatchCampaigns, 12);

			ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point1, 12, 3, BatchCampaignsRepeatingSchedule, 12);

			ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point1, 13, 3, ScheduledBatchCampaigns, 12);

			// printing Asset Data totals

			ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point1, 18, 3, LandingPages, 12);

			ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point1, 19, 3, Forms, 12);

			ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point1, 20, 3, Emails, 12);

			ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point1, 21, 3, Snippets, 12);

			ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point1, 22, 3, ImagesandFiles, 12);

			// printing Database Data
			;
			ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point1, 27, 3, AllPeople, 12);

			ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point1, 28, 3, MarketableLeads, 12);

			ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point1, 29, 3, PossibleDuplicates, 12);

			ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point1, 30, 3, UnsubscribedPeople, 12);

			ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point1, 31, 3, MarketingSuspended, 12);

			// Printing program data totals for single sheet

			ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point1, 40, 3, "D662FC", passData.Exceldata("Asset Data"),
					12);

			ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point1, 41, 3, tags, 12);

			ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point1, 42, 3, Models, 12);

		}

		ArDataPoint2(AR_Data_Point2, workbookoutput);
		workbookoutput.getCreationHelper().createFormulaEvaluator().evaluateAll();
		FileOutputStream out = new FileOutputStream(report);
		workbookoutput.write(out);
		logger.info("Printing done in AR Data Sheet 1 and 2");
		logger.info("AR report ready for view can view in Report folder");
		out.close();
		

	}

	public static void ArDataPoint2(XSSFSheet AR_Data_Point2, XSSFWorkbook workbookoutput) throws Exception {
		try {

			
			ARDataPoint2.IntrestingMomentAR(AR_Data_Point2, workbookoutput);
			// Lead scoring section
			ARDataPoint2.leadAR( AR_Data_Point2,  workbookoutput);
			
			//Printing Event  Data 
			
			ARDataPoint2.EventCampionAR(AR_Data_Point2, workbookoutput);
			
			// printing Nurture data
			ARDataPoint2.SegmentationAR(AR_Data_Point2, workbookoutput);
			
			// Printing Nurture Library
			ARDataPoint2.NurtureCampaigns(AR_Data_Point2, workbookoutput);
			
			// Printing Program Library
			ARDataPoint2.ProgrameLibraryAR(AR_Data_Point2, workbookoutput);
			
			// Printing ProgramData 
			ARDataPoint2.ProgramDataAR(AR_Data_Point2, workbookoutput);
			
			// Printing integrations
			ARDataPoint2.IntegrationAR(AR_Data_Point2, workbookoutput);
			
			// printing Revenue Models
			ARDataPoint2.RevenueModelsAR(AR_Data_Point2, workbookoutput);
			
			//Printing Web Personalization
			ARDataPoint2.Web_Personalisation(AR_Data_Point2, workbookoutput);
			
			ARDataPoint2.AddtionalProductAR(AR_Data_Point2, workbookoutput);
			


		} catch (Exception e) {
			// TODO: handle exception
		}
	}
}
