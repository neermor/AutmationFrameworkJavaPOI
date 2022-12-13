package com.marketo.qa.utility;

import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.Date;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelARReport extends converting {

	static converting obj = new converting();

	static String ExcelPath = System.getProperty("user.dir") + "//Reports//";
	static String fileName = new SimpleDateFormat("dd_MM_yy_HH_mm").format(new Date());

	public static void ARReport() throws Exception {
		
		converting.localtest();

		String AR = System.getProperty("user.dir") + "//Config//AR_Temp.xlsx";
		String report = System.getProperty("user.dir") + "//Reports//" + passData.Exceldata("Account Name") + "_"
				+ fileName + ".xlsx";
		// FileInputStream FileInputStream = new FileInputStream(report);

		// Create a blank sheet
		XSSFWorkbook workbookinput = new XSSFWorkbook(AR);
		XSSFWorkbook workbookoutput = workbookinput;

		int WorkSpace_count = Integer.parseInt(passData.Exceldata("Total WorkSpace"));

		XSSFSheet AR_Data_Point1 = workbookoutput.getSheet("AR Data Points 1");
		XSSFSheet AR_Data_Point2 = workbookoutput.getSheet("AR Data Points 2");

		// Adding prefix name
		ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point1, 1, 1, "FFFFFF", "Requested By CSM", 11);
		// adjusting column width for workspace

		AR_Data_Point1.setColumnWidth(3 + WorkSpace_count, 25 * 256);
		AR_Data_Point1.setColumnWidth(5 + WorkSpace_count, 25 * 1000);

		// inserting tag Images
		int row = 47;
		for (int i = 1; i <= 1; i++) {

			try {
				ExcelStyling.pestImg(workbookoutput, AR_Data_Point1, 2, row, passData.FetchScreenshot("Tags" + i));
				row = row + 22;

			} catch (Exception e) {
				// TODO: handle exception
			}

		}

		// printing data in sheet
		if (WorkSpace_count >= 1) {
			
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
				}
				int coll = 3;
				int j = 2;
				int k = 7;

				// Printing workspaces names
				for (int i = 1; i <= WorkSpace_count; i++) {
					
						ExcelStyling.mergeAndCenter(workbookoutput, AR_Data_Point1, "D662FC", "AR Data Points", 1, 2, 2,
								3 + WorkSpace_count, true, 25);

						ExcelStyling.mergeAndCenter(workbookoutput, AR_Data_Point1, "B541DF",
								"instance acct prefix: " + passData.Exceldata("Account Name"), 3, 3, 2,
								3 + WorkSpace_count, false, 12);

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
					ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point1, 8, 3 + WorkSpace_count, passData.Exceldata("All Batch Campaigns"),
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
				// Printing Program Data
				try {
					ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point1, 41, 3, passData.Exceldata("Tags"), 12);
				} catch (Exception e) {
					// TODO: handle exception
				}
				try {
					ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point1, 42, 3, passData.Exceldata("Models"), 12);
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
		out.close();
		

	}

	public static void ArDataPoint2(XSSFSheet AR_Data_Point2, XSSFWorkbook workbookoutput) throws Exception {
		try {
			int InterestingMoment = Integer.parseInt(passData.Exceldata("Interesting Moment"));

			if (InterestingMoment == 0) {
				ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point2, 6, 3, InterestingMoment, 12);

				ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point2, 6, 4, ARStringData.Intresting_moment, 12);

			} else {
				ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point2, 6, 3, InterestingMoment, 12);
				try {
					ExcelStyling.pestImg(workbookoutput, AR_Data_Point2, 11, 5,
							passData.FetchScreenshot("Interesting Moment"));

				} catch (Exception e) {
					// TODO: handle exception
				}
			}

			// Lead scoring section
			int lead = Integer.parseInt(passData.Exceldata("Change Score"));

			ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point2, 20, 3, lead, 12);

			if (lead == 0) {
				ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point2, 20, 6, ARStringData.No_lead, 12);

			} else if (lead <= 5) {
				ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point2, 20, 8, ARStringData.No_leadImprovement, 12);
			}

			// printing EventCampion data
			int EventCampion = Integer.parseInt(passData.Exceldata("Event Programs"));

			if (EventCampion == 0) {
				ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point2, 37, 5,
						passData.Exceldata("Account Name") + ARStringData.No_EventCampions, 12);

			} else if (EventCampion <= 1) {
				ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point2, 37, 7,
						passData.Exceldata("Account Name") + ARStringData.EventCampionsImprovment, 12);

				ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point2, 37, 3, "yes", 14);
				try {
					ExcelStyling.pestImg(workbookoutput, AR_Data_Point2, 11, 36, passData.FetchScreenshot("Event"));

				} catch (Exception e) {
					// TODO: handle exception
				}
			} else {

				ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point2, 37, 3, "yes", 14);

				try {
					ExcelStyling.pestImg(workbookoutput, AR_Data_Point2, 11, 36, passData.FetchScreenshot("Event"));

				} catch (Exception e) {
					// TODO: handle exception
				}
			}

			// printing Nurture data
			int NurtureCampaigns = Integer.parseInt(passData.Exceldata("Nurture campaigns"));

			if (NurtureCampaigns == 0) {
				ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point2, 47, 5,
						passData.Exceldata("Account Name") + ARStringData.NoNurtureCampaigns, 12);

			} else if (NurtureCampaigns <= 1) {
				ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point2, 47, 7,
						passData.Exceldata("Account Name") + ARStringData.NurtureCampaignsImprovment, 12);

				ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point2, 47, 3, "yes", 14);
				try {
					ExcelStyling.pestImg(workbookoutput, AR_Data_Point2, 11, 46,
							passData.FetchScreenshot("Nurture campaigns"));

				} catch (Exception e) {
					// TODO: handle exception
				}

			} else {

				ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point2, 47, 3, "yes", 14);
				try {
					ExcelStyling.pestImg(workbookoutput, AR_Data_Point2, 11, 46,
							passData.FetchScreenshot("Nurture campaigns"));

				} catch (Exception e) {
					// TODO: handle exception
				}

			}

			// printing Segmentation data

			int Segmentation = Integer.parseInt(passData.Exceldata("Segmentations"));

			if (Segmentation == 0) {
				ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point2, 58, 5,
						passData.Exceldata("Account Name") + ARStringData.NoSegmentation, 12);

			} else {
				ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point2, 58, 3, "Yes", 12);
			}
			int coll = 9;
			for (int i = 1; i < 3; i++)
				try {
					ExcelStyling.pestImg(workbookoutput, AR_Data_Point2, coll, 58,
							passData.FetchScreenshot("Segmentations" + i));
					coll = coll + 3;

				} catch (Exception e) {
					// TODO: handle exception
				}

			// Printing Program Library

			String ProgrameLibrary = passData.Exceldata("Programe Library");

			if (ProgrameLibrary.equalsIgnoreCase("true")) {

				ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point2, 72, 3, "Yes", 12);

			} else {
				ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point2, 72, 5, ARStringData.NoProgrameLibrary, 12);

			}

			// Printing Data Management

			// String DataManagement=passData.Exceldata("Change Data Value");

			if (DataManagement > 0) {

				ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point2, 80, 3, "Yes", 12);

			} else {
				ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point2, 80, 5,
						passData.Exceldata("Account Name") + ARStringData.NoDataManagement, 12);

			}

			// Printing integrations

			// int integrations=passData.Exceldata("Integration");

			if (Integration > 0) {

				ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point2, 92, 3, "Yes", 12);
				try {
					ExcelStyling.pestImg(workbookoutput, AR_Data_Point2, 10, 90,
							passData.FetchScreenshot("Integration"));

				} catch (Exception e) {
					// TODO: handle exception
				}

			} else {
				ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point2, 92, 5, ARStringData.NoIntegration, 12);

			}

			// printing Revenue Models
			int RevenueModels = Integer.parseInt(passData.Exceldata("Models"));

			if (RevenueModels == 0) {
				ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point2, 98, 5,
						passData.Exceldata("Account Name") + ARStringData.NoModel, 12);

			} else if (RevenueModels <= 1) {
				ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point2, 98, 7,
						passData.Exceldata("Account Name") + ARStringData.ModelsImprovment, 12);

				ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point2, 98, 3, "yes", 14);
			} else {

				ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point2, 98, 3, "yes", 14);
			}
			int WPCount = Integer.parseInt(passData.Exceldata("Total WorkSpace"));
			int j = 1;

			for (int i = 1; i <= WPCount; i++) {
				String data = "Approved Models" + j++;
				try {
					ExcelStyling.pestImg(workbookoutput, AR_Data_Point2, coll, 98,
							passData.FetchScreenshot(passData.Exceldata(data, 2)));
					coll = coll + 3;

				} catch (Exception e) {
					// TODO: handle exception
				}

				int ModelsCount = Integer.parseInt(passData.Exceldata(data));
				ModelsCount = ModelsCount + 3;
				coll = 11;
				for (int k = 3; k <= ModelsCount; k++) {
					try {

						ExcelStyling.pestImg(workbookoutput, AR_Data_Point2, coll, 102,
								passData.FetchScreenshotForApprovedModels(passData.Exceldata(data, k)));
						coll = coll + 7;
					} catch (Exception e) {
						// TODO: handle exception
					}

				}
			}

			// Additional Product WebPersonalization
			if (passData.Exceldata("Web Personalization").equalsIgnoreCase("true")) {

				ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point2, 110, 3, "Yes", 12);

				try {
					ExcelStyling.pestImg(workbookoutput, AR_Data_Point2, 2, 120,
							passData.FetchScreenshot("Top Campaigns_Default"));
					ExcelStyling.pestImg(workbookoutput, AR_Data_Point2, 5, 120,
							passData.FetchScreenshot("Top Content_Default"));
					ExcelStyling.pestImg(workbookoutput, AR_Data_Point2, 8, 120,
							passData.FetchScreenshot("Top Industries_Default"));
					ExcelStyling.pestImg(workbookoutput, AR_Data_Point2, 7, 130,
							passData.FetchScreenshot("Total Organizations_Top OrganizationsDefault"));

				} catch (Exception e) {
					// TODO: handle exception
				}
				if (TotalOrganizations <= 5) {
					ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point2, 110, 7,
							ARStringData.WebPersonalizationImprovment, 12);

				}

			} else if (passData.Exceldata("Web Personalization").equalsIgnoreCase("false")) {
				ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point2, 110, 5, ARStringData.NoWebPersonalization, 12);
			}

			// TAM data printing

			String TAM = passData.Exceldata("Target Account Management");

			if (TAM.equalsIgnoreCase("true")) {

				ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point2, 112, 3, "Yes", 12);
				ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point2, 141, 4, NamedAccounts, 12);
				ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point2, 143, 4, OpenOpportunities, 12);
				ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point2, 145, 4, Pipeline, 12);

				try {
					ExcelStyling.pestImg(workbookoutput, AR_Data_Point2, 2, 149, passData.FetchScreenshot("Overview"));

				} catch (Exception e) {
					// TODO: handle exception
				}

				if (NamedAccounts <= 5) {
					ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point2, 112, 7, ARStringData.TMAImprovment, 12);
				}

				try {
					ExcelStyling.pestImg(workbookoutput, AR_Data_Point2, 2, 149, passData.FetchScreenshot("Overview"));

				} catch (Exception e) {
					// TODO: handle exception
				}

			} else if (TAM.equalsIgnoreCase("false")) {
				ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point2, 112, 5, ARStringData.NoTAM, 12);
			}

			// Predictive Content data printing

			String PredictiveContent = passData.Exceldata("Predictive Content");

			if (PredictiveContent.equalsIgnoreCase("true")) {

				ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point2, 114, 3, "Yes", 12);
				if (TotalContent <= 5) {
					ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point2, 114, 7,
							ARStringData.PredictiveContentImprovment, 12);
				}

				try {
					ExcelStyling.pestImg(workbookoutput, AR_Data_Point2, 2, 177,
							passData.FetchScreenshot("Predictive Content"));

				} catch (Exception e) {
					// TODO: handle exception
				}
			}

			else if (PredictiveContent.equalsIgnoreCase("false")) {
				ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point2, 114, 5, ARStringData.NoPredictiveContent, 12);
			}

		} catch (Exception e) {
			// TODO: handle exception
		}
	}
}
