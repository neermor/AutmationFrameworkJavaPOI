package com.marketo.qa.utility;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ARDataPoint2 {
	
	//printing AR data Point 2 Sheet

	public static void IntrestingMomentAR(XSSFSheet AR_Data_Point2, XSSFWorkbook workbookoutput) throws Exception {
		try {
			int InterestingMoment = Integer.parseInt(passData.Exceldata("Interesting Moment"));

			if (InterestingMoment == 0) {
				ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point2, 6, 3, InterestingMoment, 12);

				ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point2, 6, 4, ARStringData.Intresting_moment, 12);

			} else {
				ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point2, 6, 3, InterestingMoment, 12);
				try {
					ExcelStyling.pestImg(workbookoutput, AR_Data_Point2, 11, 5,
							passData.FetchScreenshot("Interesting Moment"),0.6);

				} catch (Exception e) {
					// TODO: handle exception
				}
			}
		} catch (Exception e) {
			// TODO: handle exception
		}
	}

	public static void leadAR(XSSFSheet AR_Data_Point2, XSSFWorkbook workbookoutput) throws Exception {
		try {
			int lead = Integer.parseInt(passData.Exceldata("Change Score"));

			ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point2, 20, 3, lead, 12);

			if (lead == 0) {
				ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point2, 20, 6, ARStringData.No_lead, 12);

			} else if (lead <= 5) {
				ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point2, 20, 8, ARStringData.No_leadImprovement, 12);
			}

		} catch (Exception e) {
			// TODO: handle exception
		}
	}

	public static void EventCampionAR(XSSFSheet AR_Data_Point2, XSSFWorkbook workbookoutput) throws Exception {
		try {
		int EventCampion = Integer.parseInt(passData.Exceldata("Event Programs"));

		if (EventCampion == 0) {
			ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point2, 37, 5,
					passData.Exceldata("Account Name") + ARStringData.No_EventCampions, 12);

		} else if (EventCampion <= 1) {
			ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point2, 37, 7,
					passData.Exceldata("Account Name") + ARStringData.EventCampionsImprovment, 12);

			ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point2, 37, 3, "yes", 14);
			try {
				ExcelStyling.pestImg(workbookoutput, AR_Data_Point2, 11, 36, passData.FetchScreenshot("Event"),0.7);

			} catch (Exception e) {
				// TODO: handle exception
			}
		} else {

			ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point2, 37, 3, "yes", 14);

			try {
				ExcelStyling.pestImg(workbookoutput, AR_Data_Point2, 11, 36, passData.FetchScreenshot("Event"),0.7);

			} catch (Exception e) {
				// TODO: handle exception
			}
		}
		}catch (Exception e) {
			// TODO: handle exception
		}

	}

	public static void NurtureCampaigns(XSSFSheet AR_Data_Point2, XSSFWorkbook workbookoutput) throws Exception
	{
		try {
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
							passData.FetchScreenshot("Nurture campaigns"),0.7);

				} catch (Exception e) {
					// TODO: handle exception
				}

			} else {

				ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point2, 47, 3, "yes", 14);
				try {
					ExcelStyling.pestImg(workbookoutput, AR_Data_Point2, 11, 46,
							passData.FetchScreenshot("Nurture campaigns"),0.7);

				} catch (Exception e) {
					// TODO: handle exception
				}

			}
			
			
		}
		catch (Exception e) {
			// TODO: handle exception
		}
	}

	public static void SegmentationAR(XSSFSheet AR_Data_Point2, XSSFWorkbook workbookoutput) throws Exception{
		try {
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
							passData.FetchScreenshot("Segmentations" + i),0.7);
					coll = coll + 3;

				} catch (Exception e) {
					// TODO: handle exception
				}
			
		}
		catch (Exception e) {
			// TODO: handle exception
		}
	}

	public static void ProgrameLibraryAR(XSSFSheet AR_Data_Point2, XSSFWorkbook workbookoutput) throws Exception{
		try {
			

			String ProgrameLibrary = passData.Exceldata("Programe Library");

			if (ProgrameLibrary.equalsIgnoreCase("true")) {

				ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point2, 72, 3, "Yes", 12);

			} else {
				ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point2, 72, 5, ARStringData.NoProgrameLibrary, 12);

			}
		}
		catch (Exception e) {
			// TODO: handle exception
		}
		
	}

	public static void ProgramDataAR(XSSFSheet AR_Data_Point2, XSSFWorkbook workbookoutput) throws Exception{
		try {
			
			int DataManagement=Integer.parseInt(passData.Exceldata("Change Data Value"));

			if (DataManagement > 0) {

				ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point2, 80, 3, "Yes", 12);

			} else {
				ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point2, 80, 5,
						passData.Exceldata("Account Name") + ARStringData.NoDataManagement, 12);

			}

			
		}
		catch (Exception e) {
			// TODO: handle exception
		}
	}

	public static void IntegrationAR(XSSFSheet AR_Data_Point2, XSSFWorkbook workbookoutput) throws Exception{
		try {
			int Integration= Integer.parseInt(passData.Exceldata("Integration"));
			if (Integration > 0) {

				ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point2, 92, 3, "Yes", 12);
				try {
					ExcelStyling.pestImg(workbookoutput, AR_Data_Point2, 10, 90,
							passData.FetchScreenshot("Integration"),0.7);

				} catch (Exception e) {
					// TODO: handle exception
				}

			} else {
				ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point2, 92, 5, ARStringData.NoIntegration, 12);

			}
		}
		catch (Exception e) {
			// TODO: handle exception
		}
	}

	public static void RevenueModelsAR(XSSFSheet AR_Data_Point2, XSSFWorkbook workbookoutput) throws Exception {
		try {
			
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
			int coll= 10;
			for (int i = 1; i <= WPCount; i++) {
				String data = "Approved Models" + j++;
				try {
					ExcelStyling.pestImg(workbookoutput, AR_Data_Point2, coll, 98,
							passData.FetchScreenshot(passData.Exceldata(data, 2)),0.7);
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
								passData.FetchScreenshotForApprovedModels(passData.Exceldata(data, k)),0.7);
						coll = coll + 7;
					} catch (Exception e) {
						// TODO: handle exception
					}

				}
			}

		}
		catch (Exception e) {
			// TODO: handle exception
		}
	}

	public static void Web_Personalisation(XSSFSheet AR_Data_Point2, XSSFWorkbook workbookoutput) throws Exception {
		
			if (passData.Exceldata("Web Personalization").equalsIgnoreCase("true")) {

				ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point2, 110, 3, "Yes", 12);

			
					ExcelStyling.pestImg(workbookoutput, AR_Data_Point2, 2, 120,
							passData.FetchScreenshot("Top Campaigns_Default"),0.6);
					ExcelStyling.pestImg(workbookoutput, AR_Data_Point2, 5, 120,
							passData.FetchScreenshot("Top Content_Default"),0.6);
					ExcelStyling.pestImg(workbookoutput, AR_Data_Point2, 8, 120,
							passData.FetchScreenshot("Top Industries_Default"),0.6);
					ExcelStyling.pestImg(workbookoutput, AR_Data_Point2, 7, 130,
							passData.FetchScreenshot("Total Organizations_Top OrganizationsDefault"),0.6);

				
				
					int TotalOrganizations=Integer.parseInt(passData.Exceldata("Total Organizations"));
						if (TotalOrganizations <= 5) {
							ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point2, 110, 7,
									ARStringData.WebPersonalizationImprovment, 12);
		
						}
					 
			}
		
				
				else if (passData.Exceldata("Web Personalization").equalsIgnoreCase("false")) {
					ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point2, 110, 5, ARStringData.NoWebPersonalization, 12);
				}
			
			}
	
	
	public static void AddtionalProductAR(XSSFSheet AR_Data_Point2, XSSFWorkbook workbookoutput) throws Exception{
		
		

			// TAM data printing
				
			String TAM = passData.Exceldata("Target Account Management");

			if (TAM.equalsIgnoreCase("true")) {
				try {
				ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point2, 112, 3, "Yes", 12);
				int	NamedAccounts= Integer.parseInt(passData.Exceldata("Named Accounts"));
				ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point2, 141, 4, NamedAccounts, 12);
				int Pipeline= Integer.parseInt(passData.Exceldata("Pipeline"));	
				int OpenOpportunities= Integer.parseInt(passData.Exceldata("Open Opportunities"));
				ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point2, 143, 4, OpenOpportunities, 12);
				ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point2, 145, 4, Pipeline, 12);
				}
				catch (Exception e) {
					// TODO: handle exception
				}
				try {
					ExcelStyling.pestImg(workbookoutput, AR_Data_Point2, 2, 149, passData.FetchScreenshot("Overview"),0.6);

				} catch (Exception e) {
					// TODO: handle exception
				} 
				
				try {
					int	NamedAccounts= Integer.parseInt(passData.Exceldata("Named Accounts"));
				if (NamedAccounts <= 5) {
					ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point2, 112, 7, ARStringData.TMAImprovment, 12);
				}

				try {
					ExcelStyling.pestImg(workbookoutput, AR_Data_Point2, 2, 149, passData.FetchScreenshot("Overview"),0.6);

				} catch (Exception e) {
					// TODO: handle exception
				}
				}
				catch (Exception e) {
					// TODO: handle exception
				}

			} else if (TAM.equalsIgnoreCase("false")) {
				ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point2, 112, 5, ARStringData.NoTAM, 12);
			}

			// Predictive Content data printing
			
			String PredictiveContent = passData.Exceldata("Predictive Content");

			if (PredictiveContent.equalsIgnoreCase("true")) {
				
				try {
				int TotalContent= Integer.parseInt(passData.Exceldata("Total Content"));
				ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point2, 114, 3, "Yes", 12);
				try {
				if (TotalContent <= 5) {
					ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point2, 114, 7,
							ARStringData.PredictiveContentImprovment, 12);
				}
				}
				catch (Exception e) {
					// TODO: handle exception
				}

				try {
					ExcelStyling.pestImg(workbookoutput, AR_Data_Point2, 2, 177,
							passData.FetchScreenshot("Predictive Content"),0.6);

				} catch (Exception e) {
					// TODO: handle exception
				}
				}catch (Exception e) {
					// TODO: handle exception
				}
			}
			
			else if (PredictiveContent.equalsIgnoreCase("false")) {
				ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point2, 114, 5, ARStringData.NoPredictiveContent, 12);
			}
			
			
			//email insights
			
				if(PredictiveContent.equalsIgnoreCase("true"))
				{
					ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point2, 222, 3, "Yes", 12);
					try {
						ExcelStyling.pestImg(workbookoutput, AR_Data_Point2, 2, 224, passData.FetchScreenshot("Email Insights"),0.8);

					} catch (Exception e) {
						// TODO: handle exception
					}
					
				}
				else {
					ExcelStyling.WriteExcel(workbookoutput, AR_Data_Point2, 222, 5, "Client has no  Email Insights In subscription", 12);
				}
			
		
	}
	}

