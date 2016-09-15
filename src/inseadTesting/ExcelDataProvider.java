package inseadTesting;

import java.io.File;
import java.io.FileInputStream;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.Assert;
import org.testng.annotations.DataProvider;

// --------------------------------------------------------------------------------------------------------
// Data provider that reads data from an Excel file and returns results

public class ExcelDataProvider {

	// Data will be read from an Excel file, cached and then returned whenever required by test methods
	private static Object[][] mCachedData = null;

	// --------------------------------------------------------------------------------------------------------
	// The data provider method that will be called by test methods
	@DataProvider(name = "excelData")
	public static Object[][] getData()
	{
		if (null == mCachedData) 
		{
			mCachedData = readExcelFile();
		}
		Assert.assertTrue(mCachedData != null, "Unable to read data from Excel file");
		return mCachedData;
	}
	
	// --------------------------------------------------------------------------------------------------------
	// Reads an excel file and creates test data from it
	private static Object[][] readExcelFile()
	{
		List<MyInseadUser> myInseadUsers = new ArrayList<MyInseadUser>();
		TestParameters testParameters = new TestParameters();
		String executionPath = Paths.get("").toAbsolutePath().toString();
		String excelFullFilename = executionPath + "/src/inseadTesting/" + TestConstants.mTestDataExcelFile;
		try {
			FileInputStream fStream = new FileInputStream(new File(excelFullFilename));

			//Get the workbook instance for XLS file 
			Workbook workbook = new XSSFWorkbook(fStream);

			//Get first sheet from the workbook
			Sheet sheet = workbook.getSheetAt(0);

			// Iterate over all available rows
			DataFormatter formatter = new DataFormatter(); //creating formatter using the default locale
			int numRows = sheet.getLastRowNum();
			for (int iRow = 0; iRow < numRows; ++iRow)
			{
				Row row = sheet.getRow(iRow);
				if (null == row)
				{
					continue;
				}
				Cell cell = row.getCell(0);
				if (null == cell)
				{
					continue;
				}
				String firstRowCellText = formatter.formatCellValue(cell);
				if (firstRowCellText != null)
				{
					// This cell starts with a "#". Dispatch it to the appropriate reader.
					if (firstRowCellText.compareTo("#MyINSEAD test site") == 0)
					{
						populateMyInseadParameters(sheet, iRow, testParameters);
					}
					else if (firstRowCellText.compareTo("#Peoplesoft test site") == 0)
					{
						populatePeoplesoftParameters(sheet, iRow, testParameters);
					}
					else if (firstRowCellText.compareTo("#Mailchimp credentials") == 0)
					{
						populateMailChimpParameters(sheet, iRow, testParameters);
					}
					else if (firstRowCellText.compareTo("#Roles") == 0)
					{
						populateRoleParameters(sheet, iRow, testParameters);
					}
					else if (firstRowCellText.compareTo("#USER") == 0)
					{
						MyInseadUser user = new MyInseadUser();
						populateMyInseadUserParameters(sheet, iRow, user);
						myInseadUsers.add(user);
					}
					else if (firstRowCellText.compareTo("#FUNC326001_2_Update_Business_Phone_in_MyINSEAD") == 0)
					{
						testParameters.mFUNC326001_2_BusinessPhoneCountryCode = formatter.formatCellValue(sheet.getRow(iRow + 1).getCell(2));
						testParameters.mFUNC326001_2_BusinessPhone = formatter.formatCellValue(sheet.getRow(iRow + 1).getCell(4));
					}
					else if (firstRowCellText.compareTo("#TC_DirectorySearch_ByFirstname") == 0)
					{
						testParameters.mTC_DirectorySearch_ByFirstName_Name = formatter.formatCellValue(sheet.getRow(iRow + 1).getCell(2));
						testParameters.mTC_DirectorySearch_ByFirstName_Results = formatter.formatCellValue(sheet.getRow(iRow + 2).getCell(2));
					}
					else if (firstRowCellText.compareTo("#TC_MailChimp_NonMember") == 0)
					{
						testParameters.mTC_MailChimp_NonMember_EMPLID = formatter.formatCellValue(sheet.getRow(iRow + 1).getCell(1));
						testParameters.mTC_MailChimp_NonMember_List = formatter.formatCellValue(sheet.getRow(iRow + 2).getCell(1));
					}
					else if (firstRowCellText.compareTo("#TC_PreferredFlagLogic") == 0)
					{
						populatePrefFlagLogicParameters(sheet, iRow, testParameters);
					}
					
					else if (firstRowCellText.compareTo("#EMBAStudent") == 0)
					{
						testParameters.eMBAStudent1 = formatter.formatCellValue(sheet.getRow(iRow + 1).getCell(1));
						testParameters.embaStudent1Upn = formatter.formatCellValue(sheet.getRow(iRow + 1).getCell(2));
						
						testParameters.eMBAStudent2 = formatter.formatCellValue(sheet.getRow(iRow + 2).getCell(1));
						testParameters.embaStudent2Upn = formatter.formatCellValue(sheet.getRow(iRow + 2).getCell(2));
						
						testParameters.eMBAStudent3 = formatter.formatCellValue(sheet.getRow(iRow + 3).getCell(1));
						testParameters.embaStudent3Upn = formatter.formatCellValue(sheet.getRow(iRow + 3).getCell(2));
						
						testParameters.eMBAStudent4 = formatter.formatCellValue(sheet.getRow(iRow + 4).getCell(1));
						testParameters.embaStudent4Upn = formatter.formatCellValue(sheet.getRow(iRow + 4).getCell(2));
						
						testParameters.eMBAStudent5 = formatter.formatCellValue(sheet.getRow(iRow + 5).getCell(1));
						testParameters.embaStudent5Upn = formatter.formatCellValue(sheet.getRow(iRow + 5).getCell(2));

						testParameters.embaStudentSGP = formatter.formatCellValue(sheet.getRow(iRow + 6).getCell(1));
						testParameters.embaStudentSGPUPN = formatter.formatCellValue(sheet.getRow(iRow + 6).getCell(2));
						
						
					}
					else if (firstRowCellText.compareTo("#EMBAClassNbrs") == 0)
					{
						testParameters.classNbrforEmbaStu2Cor = formatter.formatCellValue(sheet.getRow(iRow + 1).getCell(1));	
						testParameters.corCourseNamewithSectionStu2 =formatter.formatCellValue(sheet.getRow(iRow + 1).getCell(2));
						testParameters.classNbrforEmbaStu2Ele =formatter.formatCellValue(sheet.getRow(iRow + 2).getCell(1));
						testParameters.eleCourseNamewithSectionStu2 =formatter.formatCellValue(sheet.getRow(iRow + 2).getCell(2));
						testParameters.classNbrforEmbaStu2Kmc =formatter.formatCellValue(sheet.getRow(iRow + 3).getCell(1));
						testParameters.kmcCourseNamewithSectionStu2 =formatter.formatCellValue(sheet.getRow(iRow + 3).getCell(2));
						testParameters.classNbrforEmbaStu2Pro =formatter.formatCellValue(sheet.getRow(iRow + 4).getCell(1));
						testParameters.proCourseNamewithSectionStu2 =formatter.formatCellValue(sheet.getRow(iRow + 4).getCell(2));
						
						
						testParameters.classNbrforEmbaStu3Cor = formatter.formatCellValue(sheet.getRow(iRow + 5).getCell(1));	
						testParameters.corCourseNamewithSectionStu3 =formatter.formatCellValue(sheet.getRow(iRow + 5).getCell(2));
						testParameters.classNbrforEmbaStu3Ele =formatter.formatCellValue(sheet.getRow(iRow + 6).getCell(1));
						testParameters.eleCourseNamewithSectionStu3 =formatter.formatCellValue(sheet.getRow(iRow + 6).getCell(2));
						testParameters.classNbrforEmbaStu3Kmc =formatter.formatCellValue(sheet.getRow(iRow + 7).getCell(1));
						testParameters.kmcCourseNamewithSectionStu3 =formatter.formatCellValue(sheet.getRow(iRow + 7).getCell(2));
						testParameters.classNbrforEmbaStu3Pro =formatter.formatCellValue(sheet.getRow(iRow + 8).getCell(1));
						testParameters.proCourseNamewithSectionStu3 =formatter.formatCellValue(sheet.getRow(iRow + 8).getCell(2));
						
						testParameters.classNbrforEmbaStu4Cor = formatter.formatCellValue(sheet.getRow(iRow + 9).getCell(1));	
						testParameters.corCourseNamewithSectionStu4 =formatter.formatCellValue(sheet.getRow(iRow + 9).getCell(2));
						testParameters.classNbrforEmbaStu4Ele =formatter.formatCellValue(sheet.getRow(iRow + 10).getCell(1));
						testParameters.eleCourseNamewithSectionStu4 =formatter.formatCellValue(sheet.getRow(iRow + 10).getCell(2));
						testParameters.classNbrforEmbaStu4Kmc =formatter.formatCellValue(sheet.getRow(iRow + 11).getCell(1));
						testParameters.kmcCourseNamewithSectionStu4 =formatter.formatCellValue(sheet.getRow(iRow + 11).getCell(2));
						testParameters.classNbrforEmbaStu4Pro =formatter.formatCellValue(sheet.getRow(iRow + 12).getCell(1));
						testParameters.proCourseNamewithSectionStu4 =formatter.formatCellValue(sheet.getRow(iRow + 12).getCell(2));
						
					}
					else if(firstRowCellText.compareTo("#TC_Job_Title") == 0){
						testParameters.student_or_alumni_user1_emplid =formatter.formatCellValue(sheet.getRow(iRow + 1).getCell(1));
						testParameters.student_or_alumni_user2_emplid =formatter.formatCellValue(sheet.getRow(iRow + 2).getCell(1));
						testParameters.student_or_alumni_user3_emplid =formatter.formatCellValue(sheet.getRow(iRow + 3).getCell(1));
						testParameters.student_or_alumni_user4_emplid =formatter.formatCellValue(sheet.getRow(iRow + 4).getCell(1));
						testParameters.student_or_alumni_user5_emplid =formatter.formatCellValue(sheet.getRow(iRow + 5).getCell(1));
						testParameters.student_or_alumni_user6_emplid =formatter.formatCellValue(sheet.getRow(iRow + 6).getCell(1));
						testParameters.industryForJobForm =formatter.formatCellValue(sheet.getRow(iRow + 7).getCell(1));
						testParameters.companyForJobForm =formatter.formatCellValue(sheet.getRow(iRow + 8).getCell(1));
						testParameters.countryForJobForm =formatter.formatCellValue(sheet.getRow(iRow + 9).getCell(1));
					
					}
					else if(firstRowCellText.compareTo("#LinkedIn credentials") == 0){
						
						populateLinkedInParameters(sheet, iRow, testParameters);
					
					}
					else if (firstRowCellText.compareTo("#TC_TermCampusDeletionIssue") == 0)
					{
						testParameters.termstudent = formatter.formatCellValue(sheet.getRow(iRow + 1).getCell(1));
						
					}
					else if(firstRowCellText.compareTo("#TC_Activation_MembReminder") == 0){
						testParameters.alumni_user1 = formatter.formatCellValue(sheet.getRow(iRow + 1).getCell(1));
						testParameters.alumni_user2 = formatter.formatCellValue(sheet.getRow(iRow + 2).getCell(1));
						testParameters.alumni_user3 = formatter.formatCellValue(sheet.getRow(iRow + 3).getCell(1));
						testParameters.naa_admin_for_user1_2_3_10_11_12 = formatter.formatCellValue(sheet.getRow(iRow + 4).getCell(1));
						testParameters.alumni_user4 = formatter.formatCellValue(sheet.getRow(iRow + 5).getCell(1));
						testParameters.alumni_user5 = formatter.formatCellValue(sheet.getRow(iRow + 6).getCell(1));
						testParameters.alumni_user6 = formatter.formatCellValue(sheet.getRow(iRow + 7).getCell(1));
						testParameters.naa_admin_for_iaa = formatter.formatCellValue(sheet.getRow(iRow + 8).getCell(1));
						testParameters.alumni_user7 = formatter.formatCellValue(sheet.getRow(iRow + 9).getCell(1));
						testParameters.alumni_user8 = formatter.formatCellValue(sheet.getRow(iRow + 10).getCell(1));
						testParameters.alumni_user9 = formatter.formatCellValue(sheet.getRow(iRow + 11).getCell(1));
						testParameters.naa_admin_for_user7_8_9_16_17_18 = formatter.formatCellValue(sheet.getRow(iRow + 12).getCell(1));
						testParameters.alumni_user10 = formatter.formatCellValue(sheet.getRow(iRow + 13).getCell(1));
						testParameters.alumni_user11 = formatter.formatCellValue(sheet.getRow(iRow + 14).getCell(1));
						testParameters.alumni_user12 = formatter.formatCellValue(sheet.getRow(iRow + 15).getCell(1));
						testParameters.alumni_user13 = formatter.formatCellValue(sheet.getRow(iRow + 16).getCell(1));
						testParameters.alumni_user14 = formatter.formatCellValue(sheet.getRow(iRow + 17).getCell(1));
						testParameters.alumni_user15 = formatter.formatCellValue(sheet.getRow(iRow + 18).getCell(1));
						testParameters.alumni_user16 = formatter.formatCellValue(sheet.getRow(iRow + 19).getCell(1));
						testParameters.alumni_user17 = formatter.formatCellValue(sheet.getRow(iRow + 20).getCell(1));
						testParameters.alumni_user18 = formatter.formatCellValue(sheet.getRow(iRow + 21).getCell(1));
						testParameters.alumni_user19 = formatter.formatCellValue(sheet.getRow(iRow + 22).getCell(1));
						testParameters.naa_admin_for_user_19 = formatter.formatCellValue(sheet.getRow(iRow + 23).getCell(1));
						testParameters.alumni_user20 = formatter.formatCellValue(sheet.getRow(iRow + 24).getCell(1));
						testParameters.alumni_user21 = formatter.formatCellValue(sheet.getRow(iRow + 25).getCell(1));
						testParameters.alumni_user22 = formatter.formatCellValue(sheet.getRow(iRow + 26).getCell(1));
						testParameters.naa_admin_for_user_22 = formatter.formatCellValue(sheet.getRow(iRow + 27).getCell(1));
						testParameters.alumni_user23 = formatter.formatCellValue(sheet.getRow(iRow + 28).getCell(1));
						testParameters.alumni_user24 = formatter.formatCellValue(sheet.getRow(iRow + 29).getCell(1));
						testParameters.alumni_user25 = formatter.formatCellValue(sheet.getRow(iRow + 30).getCell(1));
						testParameters.naa_admin_for_user_25 = formatter.formatCellValue(sheet.getRow(iRow + 31).getCell(1));
						testParameters.alumni_user26 = formatter.formatCellValue(sheet.getRow(iRow + 32).getCell(1));
						testParameters.alumni_user27 = formatter.formatCellValue(sheet.getRow(iRow + 33).getCell(1));
						

						testParameters.alumni_user28 = formatter.formatCellValue(sheet.getRow(iRow + 34).getCell(1));
						testParameters.alumni_user29 = formatter.formatCellValue(sheet.getRow(iRow + 35).getCell(1));
						testParameters.alumni_user30 = formatter.formatCellValue(sheet.getRow(iRow + 36).getCell(1));
						testParameters.alumni_user31 = formatter.formatCellValue(sheet.getRow(iRow + 37).getCell(1));
						testParameters.alumni_user32 = formatter.formatCellValue(sheet.getRow(iRow + 38).getCell(1));
						testParameters.alumni_user33 = formatter.formatCellValue(sheet.getRow(iRow + 39).getCell(1));
						testParameters.alumni_user34 = formatter.formatCellValue(sheet.getRow(iRow + 40).getCell(1));
						testParameters.alumni_user35 = formatter.formatCellValue(sheet.getRow(iRow + 41).getCell(1));
						testParameters.alumni_user36 = formatter.formatCellValue(sheet.getRow(iRow + 42).getCell(1));


						
					}
					else if(firstRowCellText.compareTo("#WebMail credentials") == 0){
						testParameters.webMailUserName = formatter.formatCellValue(sheet.getRow(iRow + 1).getCell(1));
						testParameters.webMailPassword = formatter.formatCellValue(sheet.getRow(iRow + 2).getCell(1));
					}
					else if(firstRowCellText.compareTo("#Reminder_task_settings") == 0){
						testParameters.reminderTaskUrl =formatter.formatCellValue(sheet.getRow(iRow + 1).getCell(1)); 
					}
					
					else if (firstRowCellText.compareTo("#USER_PREFERRED_EMAIL") == 0)
					{
						testParameters.mailchimp_member1 = formatter.formatCellValue(sheet.getRow(iRow + 1).getCell(1));
						testParameters.mailchimp_member1_email = formatter.formatCellValue(sheet.getRow(iRow + 1).getCell(2));
						testParameters.mailchimp_member2 = formatter.formatCellValue(sheet.getRow(iRow + 2).getCell(1));
						testParameters.mailchimp_member2_email = formatter.formatCellValue(sheet.getRow(iRow + 2).getCell(2));
						testParameters.mailchimp_member3 = formatter.formatCellValue(sheet.getRow(iRow + 3).getCell(1));
						testParameters.mailchimp_member3_email = formatter.formatCellValue(sheet.getRow(iRow + 3).getCell(2));
						testParameters.mailchimp_member4 = formatter.formatCellValue(sheet.getRow(iRow + 4).getCell(1));
						testParameters.mailchimp_member4_email = formatter.formatCellValue(sheet.getRow(iRow + 4).getCell(2));
						testParameters.mailchimp_member5 = formatter.formatCellValue(sheet.getRow(iRow + 5).getCell(1));
						testParameters.mailchimp_member5_email = formatter.formatCellValue(sheet.getRow(iRow + 5).getCell(2));
						testParameters.mailchimp_member6 = formatter.formatCellValue(sheet.getRow(iRow + 6).getCell(1));
						testParameters.mailchimp_member6_email = formatter.formatCellValue(sheet.getRow(iRow + 6).getCell(2));
						testParameters.mailchimp_member7 = formatter.formatCellValue(sheet.getRow(iRow + 7).getCell(1));
						testParameters.mailchimp_member7_email = formatter.formatCellValue(sheet.getRow(iRow + 7).getCell(2));
						testParameters.mailchimp_member8 = formatter.formatCellValue(sheet.getRow(iRow + 8).getCell(1));
						testParameters.mailchimp_member8_email = formatter.formatCellValue(sheet.getRow(iRow + 8).getCell(2));
						testParameters.mailchimp_member9 = formatter.formatCellValue(sheet.getRow(iRow + 9).getCell(1));
						testParameters.mailchimp_member9_email = formatter.formatCellValue(sheet.getRow(iRow + 9).getCell(2));
						testParameters.mailchimp_member10 = formatter.formatCellValue(sheet.getRow(iRow + 10).getCell(1));
						testParameters.mailchimp_member10_email = formatter.formatCellValue(sheet.getRow(iRow + 10).getCell(2));

					}
					else if (firstRowCellText.compareTo("#TC_Job_Title_Not_Mandatory") == 0){
						testParameters.stry0010503_user1 = formatter.formatCellValue(sheet.getRow(iRow + 1).getCell(1));
						testParameters.stry0010503_user2 = formatter.formatCellValue(sheet.getRow(iRow + 2).getCell(1));
					}
					
				}
			}
			// Get second sheet from the workbook
						Sheet sheet2 = workbook.getSheetAt(1);

						// Iterate over all available rows
						DataFormatter formatter2 = new DataFormatter(); // creating
																		// formatter using
																		// the default
																		// locale
						int numRows2 = sheet2.getLastRowNum();
						for (int iRow = 0; iRow < numRows2; ++iRow) {
							Row row = sheet2.getRow(iRow);
							if (null == row) {
								continue;
							}
							Cell cell = row.getCell(0);
							if (null == cell) {
								continue;
							}
							String firstRowCellText = formatter2.formatCellValue(cell);
							if (firstRowCellText != null) {
								// This cell starts with a "#". Dispatch it to the
								// appropriate reader.
								if (firstRowCellText.compareTo("#Regression") == 0) {
									populateRegressionEmplidParameters(sheet2, iRow, testParameters);
								} else if (firstRowCellText.compareTo("#AddNewPerson") == 0) {
									populateNewPersonParameters(sheet2, iRow, testParameters);
								} else if (firstRowCellText.compareTo("#ConstituentType") == 0) {
									populateConstituentTypeParameters(sheet2, iRow, testParameters);
								} else if (firstRowCellText.compareTo("#Phone") == 0) {
									populatePhoneParameters(sheet2, iRow, testParameters);
								} else if (firstRowCellText.compareTo("#Email") == 0) {
									populateEmailParameters(sheet2, iRow, testParameters);
								} else if (firstRowCellText.compareTo("#Nationality") == 0) {
									populateNationalityParameters(sheet2, iRow, testParameters);
								} else if (firstRowCellText.compareTo("#Language") == 0) {
									populateLanguageParameters(sheet2, iRow, testParameters);
								} else if (firstRowCellText.compareTo("#WorkExperience") == 0) {
									populateWorkExperienceParameters(sheet2, iRow, testParameters);
								}
							}

						}
						
						
						// Get third sheet from the workbook
						Sheet sheet3 = workbook.getSheetAt(2);

						// Iterate over all available rows
						DataFormatter formatter3 = new DataFormatter(); // creating
																		// formatter using
																		// the default
																		// locale
						int numRows3 = sheet3.getLastRowNum();
						for (int iRow = 0; iRow < numRows3; ++iRow) {
							Row row = sheet3.getRow(iRow);
							if (null == row) {
								continue;
							}
							Cell cell = row.getCell(0);
							if (null == cell) {
								continue;
							}
							String firstRowCellText = formatter3.formatCellValue(cell);
							if (firstRowCellText != null) {
								// This cell starts with a "#". Dispatch it to the
								// appropriate reader.
								 if (firstRowCellText.compareTo("#Phone") == 0) {
									 populateMYPhoneParameters(sheet3, iRow, testParameters);
								} 
							}

						}
			// Close out all the files
			workbook.close();
			fStream.close();
		} catch (Exception e) {
			System.out.println(e);
			Assert.assertTrue(false, "Unable to open excel file at " + excelFullFilename);
		}

		// Create the data provider object
		Object[][] testData = new Object[][] { new Object[myInseadUsers.size()] };
		
		for (int iUser = 0; iUser < myInseadUsers.size(); ++iUser)
		{
			testData[iUser] = new Object[] {testParameters, myInseadUsers.get(iUser)};
		}
		return testData;
	}
	
	// --------------------------------------------------------------------------------------------------------
	// Read MyInsead parameters from an excel file
	private static void populateMyInseadParameters(Sheet sheet, int startingRow, TestParameters testParameters)
	{
		// Make sure we are reading the right block of data
		String firstRowCellText = sheet.getRow(startingRow).getCell(0).getRichStringCellValue().getString();
		Assert.assertTrue(firstRowCellText.compareTo("#MyINSEAD test site") == 0);
		
		// Get MyInsead parameters
		DataFormatter formatter = new DataFormatter(); //creating formatter using the default locale
		testParameters.mURLMyInsead = formatter.formatCellValue(sheet.getRow(startingRow + 1).getCell(1));
		testParameters.mMyInseadGlobalAdminLogin = formatter.formatCellValue(sheet.getRow(startingRow + 2).getCell(1));
		testParameters.mMyInseadGlobalAdminPassword = formatter.formatCellValue(sheet.getRow(startingRow + 3).getCell(1));
	}
	
	// --------------------------------------------------------------------------------------------------------
	// Read Peoplesoft parameters from an excel file
	private static void populatePeoplesoftParameters(Sheet sheet, int startingRow, TestParameters testParameters)
	{
		// Make sure we are reading the right block of data
		String firstRowCellText = sheet.getRow(startingRow).getCell(0).getRichStringCellValue().getString();
		Assert.assertTrue(firstRowCellText.compareTo("#Peoplesoft test site") == 0);
		
		// Get Peoplesoft parameters
		DataFormatter formatter = new DataFormatter(); //creating formatter using the default locale
		testParameters.mURLPeoplesoft = formatter.formatCellValue(sheet.getRow(startingRow + 1).getCell(1));
		testParameters.mPeoplesoftLogin = formatter.formatCellValue(sheet.getRow(startingRow + 2).getCell(1));
		testParameters.mPeoplesoftPassword = formatter.formatCellValue(sheet.getRow(startingRow + 3).getCell(1));
	}
	
		// Read LinkedIn parameters from an excel file
	private static void populateLinkedInParameters(Sheet sheet, int startingRow, TestParameters testParameters)
	{
		// Make sure we are reading the right block of data
		String firstRowCellText = sheet.getRow(startingRow).getCell(0).getRichStringCellValue().getString();
		Assert.assertTrue(firstRowCellText.compareTo("#LinkedIn credentials") == 0);
		
		// Get LinkedIn parameters
		DataFormatter formatter = new DataFormatter(); //creating formatter using the default locale
		testParameters.linkedInUserName = formatter.formatCellValue(sheet.getRow(startingRow + 1).getCell(1));
		testParameters.linkedInPassword = formatter.formatCellValue(sheet.getRow(startingRow + 2).getCell(1));
		
	}
	
	// --------------------------------------------------------------------------------------------------------
	// Read MailChimp parameters from an excel file
	private static void populateMailChimpParameters(Sheet sheet, int startingRow, TestParameters testParameters)
	{
		// Make sure we are reading the right block of data
		String firstRowCellText = sheet.getRow(startingRow).getCell(0).getRichStringCellValue().getString();
		Assert.assertTrue(firstRowCellText.compareTo("#Mailchimp credentials") == 0);
		
		// Get Peoplesoft parameters
		DataFormatter formatter = new DataFormatter(); //creating formatter using the default locale
		testParameters.mMailChimpLogin = formatter.formatCellValue(sheet.getRow(startingRow + 1).getCell(1));
		testParameters.mMailChimpPassword = formatter.formatCellValue(sheet.getRow(startingRow + 2).getCell(1));
	}
	
	// --------------------------------------------------------------------------------------------------------
	// Read MyInsead user parameters from an excel file
	private static void populateMyInseadUserParameters(Sheet sheet, int startingRow, MyInseadUser user)
	{
		// Make sure we are reading the right block of data
		String firstRowCellText = sheet.getRow(startingRow).getCell(0).getRichStringCellValue().getString();
		Assert.assertTrue(firstRowCellText.compareTo("#USER") == 0);
		
		// Get MyInsead user parameters
		DataFormatter formatter = new DataFormatter(); //creating formatter using the default locale
		user.mEMPLID = formatter.formatCellValue(sheet.getRow(startingRow + 1).getCell(2));
		user.mPrimaryFirstName = formatter.formatCellValue(sheet.getRow(startingRow + 1).getCell(4));
		user.mPrimaryLastName = formatter.formatCellValue(sheet.getRow(startingRow + 1).getCell(6));
		
		user.mCompany = formatter.formatCellValue(sheet.getRow(startingRow + 3).getCell(2));
		user.mBillingEmailAddress = formatter.formatCellValue(sheet.getRow(startingRow + 3).getCell(4));
		user.mStreetAddress1 = formatter.formatCellValue(sheet.getRow(startingRow + 4).getCell(2));
		user.mStreetAddress2 = formatter.formatCellValue(sheet.getRow(startingRow + 4).getCell(4));
		user.mStreetAddress3 = formatter.formatCellValue(sheet.getRow(startingRow + 4).getCell(6));
		user.mCity = formatter.formatCellValue(sheet.getRow(startingRow + 5).getCell(2));
		user.mStateProvince = formatter.formatCellValue(sheet.getRow(startingRow + 5).getCell(4));
		user.mZipPostalCode = formatter.formatCellValue(sheet.getRow(startingRow + 5).getCell(6));
		user.mBillingCountry = formatter.formatCellValue(sheet.getRow(startingRow + 5).getCell(8));
		user.mTelephone = formatter.formatCellValue(sheet.getRow(startingRow + 6).getCell(2));
		user.mFax = formatter.formatCellValue(sheet.getRow(startingRow + 6).getCell(4));

		user.mBusinessPhoneCountryCode = formatter.formatCellValue(sheet.getRow(startingRow + 7).getCell(2));
		user.mBusinessPhone = formatter.formatCellValue(sheet.getRow(startingRow + 7).getCell(4));

		user.mLocation = formatter.formatCellValue(sheet.getRow(startingRow + 8).getCell(2));
	}
	// --------------------------------------------------------------------------------------------------------
	// Read Roles parameters from an excel file
	private static void populateRoleParameters(Sheet sheet, int startingRow, TestParameters testParameters)
	{
		// Make sure we are reading the right block of data
		String firstRowCellText = sheet.getRow(startingRow).getCell(0).getRichStringCellValue().getString();
		Assert.assertTrue(firstRowCellText.compareTo("#Roles") == 0);
		
		// Get Roles parameters
		DataFormatter formatter = new DataFormatter(); //creating formatter using the default locale
		testParameters.mRolesAlumni = formatter.formatCellValue(sheet.getRow(startingRow + 1).getCell(1));
		testParameters.mRolesStudent = formatter.formatCellValue(sheet.getRow(startingRow + 2).getCell(1));
		testParameters.mRolesStaffFaculty = formatter.formatCellValue(sheet.getRow(startingRow + 3).getCell(1));
		testParameters.mRolesEDP = formatter.formatCellValue(sheet.getRow(startingRow + 4).getCell(1));
		testParameters.mRolesMBAAdmin = formatter.formatCellValue(sheet.getRow(startingRow + 5).getCell(1));
		testParameters.mRolesGlobalAdmin = formatter.formatCellValue(sheet.getRow(startingRow + 6).getCell(1));
		testParameters.mRolesNAAAdmin = formatter.formatCellValue(sheet.getRow(startingRow + 7).getCell(1));
		testParameters.mRolesEMBAstudentByName = formatter.formatCellValue(sheet.getRow(startingRow + 8).getCell(1));
		testParameters.MIS50_STRY0010302_AffiliateViewMember =  formatter.formatCellValue(sheet.getRow(startingRow + 9).getCell(1));
		testParameters.MIS50_STRY0010302_AffiliateViewNonMember  =  formatter.formatCellValue(sheet.getRow(startingRow + 10).getCell(1));
		testParameters.MIS50_STRY0010302_AlumniMemberNAAadmin  =  formatter.formatCellValue(sheet.getRow(startingRow + 11).getCell(1));
		testParameters.MIS50_STRY0010302_AlumniNonMemberNAAadmin  =  formatter.formatCellValue(sheet.getRow(startingRow + 12).getCell(1));
		testParameters.MIS50_STRY0010302_AffiliateMemberNAAadmin  =  formatter.formatCellValue(sheet.getRow(startingRow + 13).getCell(1));
		testParameters.MIS50_STRY0010302_AffiliateNonMemberNAAadmin  =  formatter.formatCellValue(sheet.getRow(startingRow + 14).getCell(1));
		testParameters.MIS50_STRY0010302_AffiliateNonNAAadmin  =  formatter.formatCellValue(sheet.getRow(startingRow + 15).getCell(1));
		testParameters.MIS50_STRY0010302_AffiliateOtherNAAadmin  =  formatter.formatCellValue(sheet.getRow(startingRow + 16).getCell(1));
		testParameters.MIS50_STRY0010302_AffiliateOtherNonNAAadmin  =  formatter.formatCellValue(sheet.getRow(startingRow + 17).getCell(1));
		testParameters.mRolesAffiliate  =  formatter.formatCellValue(sheet.getRow(startingRow + 18).getCell(1));
		testParameters.AlumniFirstName  =  formatter.formatCellValue(sheet.getRow(startingRow + 19).getCell(1));
		testParameters.AlumniLastName  =  formatter.formatCellValue(sheet.getRow(startingRow + 20).getCell(1));
		testParameters.StudentS52  =  formatter.formatCellValue(sheet.getRow(startingRow + 21).getCell(1));
		testParameters.StudentFirstName  =  formatter.formatCellValue(sheet.getRow(startingRow + 22).getCell(1));
		testParameters.StudentLastName  =  formatter.formatCellValue(sheet.getRow(startingRow + 23).getCell(1));

		}
	// --------------------------------------------------------------------------------------------------------
	// Read Preferred Flag Logic parameters from an excel file
	private static void populatePrefFlagLogicParameters(Sheet sheet, int startingRow, TestParameters testParameters)
	{
		// Make sure we are reading the right block of data
		String firstRowCellText = sheet.getRow(startingRow).getCell(0).getRichStringCellValue().getString();
		Assert.assertTrue(firstRowCellText.compareTo("#TC_PreferredFlagLogic") == 0);
		
		int row = startingRow + 1;
		// Get parameters
		DataFormatter formatter = new DataFormatter(); //creating formatter using the default locale
		testParameters.woPreferredPhone = formatter.formatCellValue(sheet.getRow(row++).getCell(1));
		testParameters.woPreferredEmail = formatter.formatCellValue(sheet.getRow(row++).getCell(1));
		testParameters.woPreferredAddress = formatter.formatCellValue(sheet.getRow(row++).getCell(1));
		testParameters.wHomePreferredAddress = formatter.formatCellValue(sheet.getRow(row++).getCell(1));
	}
	// --------------------------------------------------------------------------------------------------------
		// Read Regression parameters from an excel file
		private static void populateRegressionEmplidParameters(Sheet sheet, int startingRow,
				TestParameters testParameters) {
			// Make sure we are reading the right block of data
			String firstRowCellText = sheet.getRow(startingRow).getCell(0).getRichStringCellValue().getString();
			Assert.assertTrue(firstRowCellText.compareTo("#Regression") == 0);

			int row = startingRow + 1;
			// Get parameters
			DataFormatter formatter = new DataFormatter(); // creating formatter
															// using the default
															// locale
			testParameters.CRUDAlumni = formatter.formatCellValue(sheet.getRow(startingRow + 1).getCell(1)).trim();
			testParameters.CRUDStudent = formatter.formatCellValue(sheet.getRow(startingRow + 2).getCell(1)).trim();
		}

		// --------------------------------------------------------------------------------------------------------
		// Read AddNewPerson parameters from an excel file
		private static void populateNewPersonParameters(Sheet sheet, int startingRow, TestParameters testParameters) {
			// Make sure we are reading the right block of data
			String firstRowCellText = sheet.getRow(startingRow).getCell(0).getRichStringCellValue().getString();
			Assert.assertTrue(firstRowCellText.compareTo("#AddNewPerson") == 0);

			int row = startingRow + 1;
			// Get parameters
			DataFormatter formatter = new DataFormatter(); // creating formatter
															// using the default
															// locale
			testParameters.AddPersonFirstNameAlumni = formatter.formatCellValue(sheet.getRow(startingRow + 1).getCell(1)).trim();
			testParameters.AddPersonLastNameAlumni = formatter.formatCellValue(sheet.getRow(startingRow + 1).getCell(2)).trim();
			testParameters.AddPersonEmailAlumni = formatter.formatCellValue(sheet.getRow(startingRow + 1).getCell(3)).trim();
			testParameters.AddPersonEmplidAlumni = formatter.formatCellValue(sheet.getRow(startingRow + 1).getCell(4))
					.trim();

			testParameters.AddPersonFirstNameStudent = formatter.formatCellValue(sheet.getRow(startingRow + 2).getCell(1)).trim();
			testParameters.AddPersonLastNameStudent = formatter.formatCellValue(sheet.getRow(startingRow + 2).getCell(2)).trim();
			testParameters.AddPersonEmailStudent = formatter.formatCellValue(sheet.getRow(startingRow + 2).getCell(3)).trim();
			testParameters.AddPersonEmplidStudent = formatter.formatCellValue(sheet.getRow(startingRow + 2).getCell(4))
					.trim();

			testParameters.AddPersonFirstNameStaff = formatter.formatCellValue(sheet.getRow(startingRow + 3).getCell(1)).trim();
			testParameters.AddPersonLastNameStaff = formatter.formatCellValue(sheet.getRow(startingRow + 3).getCell(2)).trim();
			testParameters.AddPersonEmailStaff = formatter.formatCellValue(sheet.getRow(startingRow + 3).getCell(3)).trim();
			testParameters.AddPersonEmplidStaff = formatter.formatCellValue(sheet.getRow(startingRow + 3).getCell(4))
					.trim();

			testParameters.AddPersonFirstNameFaculty = formatter.formatCellValue(sheet.getRow(startingRow + 4).getCell(1)).trim();
			testParameters.AddPersonLastNameFaculty = formatter.formatCellValue(sheet.getRow(startingRow + 4).getCell(2)).trim();
			testParameters.AddPersonEmailFaculty = formatter.formatCellValue(sheet.getRow(startingRow + 4).getCell(3)).trim();
			testParameters.AddPersonEmplidFaculty = formatter.formatCellValue(sheet.getRow(startingRow + 4).getCell(4))
					.trim();
		}
	// --------------------------------------------------------------------------------------------------------
		
	// Read Constituent Type parameters from an excel file
		private static void populateConstituentTypeParameters(Sheet sheet, int startingRow, TestParameters testParameters) {
			// Make sure we are reading the right block of data
			String firstRowCellText = sheet.getRow(startingRow).getCell(0).getRichStringCellValue().getString();
			Assert.assertTrue(firstRowCellText.compareTo("#ConstituentType") == 0);

			DataFormatter formatter = new DataFormatter(); // creating formatter
															// using the default
															// locale
			testParameters.constituentTypeAlumni = formatter.formatCellValue(sheet.getRow(startingRow + 1).getCell(0)).trim();
			testParameters.constituentTypeAlumniValue = formatter.formatCellValue(sheet.getRow(startingRow + 1).getCell(1)).trim();
			testParameters.constituentTypeAlumniEmplid = formatter.formatCellValue(sheet.getRow(startingRow + 1).getCell(2))
					.trim();
			testParameters.constituentTypeAlumniUpdate = formatter.formatCellValue(sheet.getRow(startingRow + 1).getCell(3))
					.trim();
			testParameters.constituentTypeAlumniDelete = formatter.formatCellValue(sheet.getRow(startingRow + 1).getCell(4))
					.trim();

			testParameters.constituentTypeStudent = formatter.formatCellValue(sheet.getRow(startingRow + 2).getCell(0)).trim();
			testParameters.constituentTypeStudentValue = formatter
					.formatCellValue(sheet.getRow(startingRow + 2).getCell(1)).trim();
			testParameters.constituentTypeStudentEmplid = formatter
					.formatCellValue(sheet.getRow(startingRow + 2).getCell(2)).trim();
			testParameters.constituentTypeStudentUpdate = formatter
					.formatCellValue(sheet.getRow(startingRow + 2).getCell(3)).trim();
			testParameters.constituentTypeStudentDelete = formatter
					.formatCellValue(sheet.getRow(startingRow + 2).getCell(4)).trim();

			testParameters.constituentTypeFaculty = formatter.formatCellValue(sheet.getRow(startingRow + 3).getCell(0)).trim();
			testParameters.constituentTypeFacultyValue = formatter
					.formatCellValue(sheet.getRow(startingRow + 3).getCell(1)).trim();
			testParameters.constituentTypeFacultyEmplid = formatter
					.formatCellValue(sheet.getRow(startingRow + 3).getCell(2)).trim();
			testParameters.constituentTypeFacultyUpdate = formatter
					.formatCellValue(sheet.getRow(startingRow + 3).getCell(3)).trim();
			testParameters.constituentTypeFacultyDelete = formatter
					.formatCellValue(sheet.getRow(startingRow + 3).getCell(4)).trim();

			testParameters.constituentTypeStaff = formatter.formatCellValue(sheet.getRow(startingRow + 4).getCell(0)).trim();
			testParameters.constituentTypeStaffValue = formatter.formatCellValue(sheet.getRow(startingRow + 4).getCell(1)).trim();
			testParameters.constituentTypeStaffEmplid = formatter.formatCellValue(sheet.getRow(startingRow + 4).getCell(2))
					.trim();
			testParameters.constituentTypeStaffUpdate = formatter.formatCellValue(sheet.getRow(startingRow + 4).getCell(3))
					.trim();
			testParameters.constituentTypeStaffDelete = formatter.formatCellValue(sheet.getRow(startingRow + 4).getCell(4))
					.trim();

			testParameters.constituentTypeINSEADClient = formatter
					.formatCellValue(sheet.getRow(startingRow + 5).getCell(0)).trim();
			testParameters.constituentTypeINSEADClientValue = formatter
					.formatCellValue(sheet.getRow(startingRow + 5).getCell(1)).trim();
			testParameters.constituentTypeINSEADClientEmplid = formatter
					.formatCellValue(sheet.getRow(startingRow + 5).getCell(2)).trim();
			testParameters.constituentTypeINSEADClientUpdate = formatter
					.formatCellValue(sheet.getRow(startingRow + 5).getCell(3)).trim();
			testParameters.constituentTypeINSEADClientDelete = formatter
					.formatCellValue(sheet.getRow(startingRow + 5).getCell(4)).trim();

			testParameters.constituentTypeINSEADContractor = formatter
					.formatCellValue(sheet.getRow(startingRow + 6).getCell(0)).trim();
			testParameters.constituentTypeINSEADContractorValue = formatter
					.formatCellValue(sheet.getRow(startingRow + 6).getCell(1)).trim();
			testParameters.constituentTypeINSEADContractorEmplid = formatter
					.formatCellValue(sheet.getRow(startingRow + 6).getCell(2)).trim();
			testParameters.constituentTypeINSEADContractorUpdate = formatter
					.formatCellValue(sheet.getRow(startingRow + 6).getCell(3)).trim();
			testParameters.constituentTypeINSEADContractorDelete = formatter
					.formatCellValue(sheet.getRow(startingRow + 6).getCell(4)).trim();

			testParameters.constituentTypeAffiliate = formatter.formatCellValue(sheet.getRow(startingRow + 7).getCell(0)).trim();
			testParameters.constituentTypeAffiliateValue = formatter
					.formatCellValue(sheet.getRow(startingRow + 7).getCell(1)).trim();
			testParameters.constituentTypeAffiliateEmplid = formatter
					.formatCellValue(sheet.getRow(startingRow + 7).getCell(2)).trim();
			testParameters.constituentTypeAffiliateUpdate = formatter
					.formatCellValue(sheet.getRow(startingRow + 7).getCell(3)).trim();
			testParameters.constituentTypeAffiliateDelete = formatter
					.formatCellValue(sheet.getRow(startingRow + 7).getCell(4)).trim();

			testParameters.constituentTypeParticipant = formatter.formatCellValue(sheet.getRow(startingRow + 8).getCell(0)).trim();
			testParameters.constituentTypeParticipantValue = formatter
					.formatCellValue(sheet.getRow(startingRow + 8).getCell(1)).trim();
			testParameters.constituentTypeParticipantEmplid = formatter
					.formatCellValue(sheet.getRow(startingRow + 8).getCell(2)).trim();
			testParameters.constituentTypeParticipantUpdate = formatter
					.formatCellValue(sheet.getRow(startingRow + 8).getCell(3)).trim();
			testParameters.constituentTypeParticipantDelete = formatter
					.formatCellValue(sheet.getRow(startingRow + 8).getCell(4)).trim();

			testParameters.constituentTypeExchangerStudent = formatter
					.formatCellValue(sheet.getRow(startingRow + 9).getCell(0)).trim();
			testParameters.constituentTypeExchangerStudentValue = formatter
					.formatCellValue(sheet.getRow(startingRow + 9).getCell(1)).trim();
			testParameters.constituentTypeExchangerStudentEmplid = formatter
					.formatCellValue(sheet.getRow(startingRow + 9).getCell(2)).trim();
			testParameters.constituentTypeExchangerStudentUpdate = formatter
					.formatCellValue(sheet.getRow(startingRow + 9).getCell(3)).trim();
			testParameters.constituentTypeExchangerStudentDelete = formatter
					.formatCellValue(sheet.getRow(startingRow + 9).getCell(4)).trim();

			testParameters.constituentTypePastParticipant = formatter
					.formatCellValue(sheet.getRow(startingRow + 10).getCell(0)).trim();
			testParameters.constituentTypePastParticipantValue = formatter
					.formatCellValue(sheet.getRow(startingRow + 10).getCell(1)).trim();
			testParameters.constituentTypePastParticipantEmplid = formatter
					.formatCellValue(sheet.getRow(startingRow + 10).getCell(2)).trim();
			testParameters.constituentTypePastParticipantUpdate = formatter
					.formatCellValue(sheet.getRow(startingRow + 10).getCell(3)).trim();
			testParameters.constituentTypePastParticipantDelete = formatter
					.formatCellValue(sheet.getRow(startingRow + 10).getCell(4)).trim();
		}

		// --------------------------------------------------------------------------------------------------------
		// Read Email parameters from an excel file
		private static void populateEmailParameters(Sheet sheet, int startingRow, TestParameters testParameters) {
			// Make sure we are reading the right block of data
			String firstRowCellText = sheet.getRow(startingRow).getCell(0).getRichStringCellValue().getString();
			Assert.assertTrue(firstRowCellText.compareTo("#Email") == 0);

			int row = startingRow + 1;
			// Get parameters
			DataFormatter formatter = new DataFormatter(); // creating formatter
															// using the default
															// locale
			testParameters.HomeEmailType = formatter.formatCellValue(sheet.getRow(startingRow + 1).getCell(0)).trim();
			testParameters.HomeEmailAddress = formatter.formatCellValue(sheet.getRow(startingRow + 1).getCell(1)).trim();
			testParameters.HomeEmailAddressUpdated = formatter.formatCellValue(sheet.getRow(startingRow + 1).getCell(2)).trim();
			testParameters.HomeEmailEmplidCreate = formatter.formatCellValue(sheet.getRow(startingRow + 1).getCell(3))
					.trim();
			testParameters.HomeEmailEmplidUpdate = formatter.formatCellValue(sheet.getRow(startingRow + 1).getCell(4))
					.trim();
			testParameters.HomeEmailEmplidDelete = formatter.formatCellValue(sheet.getRow(startingRow + 1).getCell(5))
					.trim();

			testParameters.BusinessEmailType = formatter.formatCellValue(sheet.getRow(startingRow + 2).getCell(0)).trim();
			testParameters.BusinessEmailAddress = formatter.formatCellValue(sheet.getRow(startingRow + 2).getCell(1)).trim();
			testParameters.BusinessEmailAddressUpdated = formatter
					.formatCellValue(sheet.getRow(startingRow + 2).getCell(2)).trim();
			testParameters.BusinessEmailEmplidCreate = formatter.formatCellValue(sheet.getRow(startingRow + 2).getCell(3))
					.trim();
			testParameters.BusinessEmailEmplidUpdate = formatter.formatCellValue(sheet.getRow(startingRow + 2).getCell(4))
					.trim();
			testParameters.BusinessEmailEmplidDelete = formatter.formatCellValue(sheet.getRow(startingRow + 2).getCell(5))
					.trim();

			testParameters.UPNEmailType = formatter.formatCellValue(sheet.getRow(startingRow + 3).getCell(0)).trim();
			testParameters.UPNEmailAddress = formatter.formatCellValue(sheet.getRow(startingRow + 3).getCell(1)).trim();
			testParameters.UPNEmailAddressUpdated = formatter.formatCellValue(sheet.getRow(startingRow + 3).getCell(2)).trim();
			testParameters.UPNEmailEmplidCreate = formatter.formatCellValue(sheet.getRow(startingRow + 3).getCell(3))
					.trim();
			testParameters.UPNEmailEmplidUpdate = formatter.formatCellValue(sheet.getRow(startingRow + 3).getCell(4))
					.trim();
			testParameters.UPNEmailEmplidDelete = formatter.formatCellValue(sheet.getRow(startingRow + 3).getCell(5))
					.trim();
		}

		// --------------------------------------------------------------------------------------------------------
		// Read Phone parameters from an excel file
		private static void populatePhoneParameters(Sheet sheet, int startingRow, TestParameters testParameters) {
			// Make sure we are reading the right block of data
			String firstRowCellText = sheet.getRow(startingRow).getCell(0).getRichStringCellValue().getString();
			Assert.assertTrue(firstRowCellText.compareTo("#Phone") == 0);

			startingRow = startingRow + 1;
			// Get parameters
			DataFormatter formatter = new DataFormatter(); // creating formatter
															// using the default
															// locale
			testParameters.HomePhoneType = formatter.formatCellValue(sheet.getRow(startingRow + 1).getCell(0)).trim();
			testParameters.HomePhoneNum = formatter.formatCellValue(sheet.getRow(startingRow + 1).getCell(1)).trim();
			testParameters.HomePhoneExt = formatter.formatCellValue(sheet.getRow(startingRow + 1).getCell(2)).trim();
			testParameters.HomePhoneCountry = formatter.formatCellValue(sheet.getRow(startingRow + 1).getCell(3)).trim();
			testParameters.HomePhoneNumUpdated = formatter.formatCellValue(sheet.getRow(startingRow + 1).getCell(4)).trim();
			testParameters.HomePhoneExtUpdated = formatter.formatCellValue(sheet.getRow(startingRow + 1).getCell(5)).trim();
			testParameters.HomePhoneCountryUpdated = formatter.formatCellValue(sheet.getRow(startingRow + 1).getCell(6)).trim();
			testParameters.HomeEmplidCreate = formatter.formatCellValue(sheet.getRow(startingRow + 1).getCell(7)).trim();
			testParameters.HomeEmplidUpdate = formatter.formatCellValue(sheet.getRow(startingRow + 1).getCell(8)).trim();
			testParameters.HomeEmplidDelete = formatter.formatCellValue(sheet.getRow(startingRow + 1).getCell(9)).trim();

			testParameters.BusinessPhoneType = formatter.formatCellValue(sheet.getRow(startingRow + 2).getCell(0)).trim();
			testParameters.BusinessPhoneNum = formatter.formatCellValue(sheet.getRow(startingRow + 2).getCell(1)).trim();
			testParameters.BusinessPhoneExt = formatter.formatCellValue(sheet.getRow(startingRow + 2).getCell(2)).trim();
			testParameters.BusinessPhoneCountry = formatter.formatCellValue(sheet.getRow(startingRow + 2).getCell(3)).trim();
			testParameters.BusinessPhoneNumUpdated = formatter.formatCellValue(sheet.getRow(startingRow + 2).getCell(4)).trim();
			testParameters.BusinessPhoneExtUpdated = formatter.formatCellValue(sheet.getRow(startingRow + 2).getCell(5)).trim();
			testParameters.BusinessPhoneCountryUpdated = formatter
					.formatCellValue(sheet.getRow(startingRow + 2).getCell(6)).trim();
			testParameters.BusinessEmplidCreate = formatter.formatCellValue(sheet.getRow(startingRow + 2).getCell(7))
					.trim();
			testParameters.BusinessEmplidUpdate = formatter.formatCellValue(sheet.getRow(startingRow + 2).getCell(8))
					.trim();
			testParameters.BusinessEmplidDelete = formatter.formatCellValue(sheet.getRow(startingRow + 2).getCell(9))
					.trim();

			testParameters.MobilePhoneType = formatter.formatCellValue(sheet.getRow(startingRow + 3).getCell(0)).trim();
			testParameters.MobilePhoneNum = formatter.formatCellValue(sheet.getRow(startingRow + 3).getCell(1)).trim();
			testParameters.MobilePhoneExt = formatter.formatCellValue(sheet.getRow(startingRow + 3).getCell(2)).trim();
			testParameters.MobilePhoneCountry = formatter.formatCellValue(sheet.getRow(startingRow + 3).getCell(3)).trim();
			testParameters.MobilePhoneNumUpdated = formatter.formatCellValue(sheet.getRow(startingRow + 3).getCell(4)).trim();
			testParameters.MobilePhoneExtUpdated = formatter.formatCellValue(sheet.getRow(startingRow + 3).getCell(5)).trim();
			testParameters.MobilePhoneCountryUpdated = formatter.formatCellValue(sheet.getRow(startingRow + 3).getCell(6)).trim();
			testParameters.MobileEmplidCreate = formatter.formatCellValue(sheet.getRow(startingRow + 3).getCell(7)).trim();
			testParameters.MobileEmplidUpdate = formatter.formatCellValue(sheet.getRow(startingRow + 3).getCell(8)).trim();
			testParameters.MobileEmplidDelete = formatter.formatCellValue(sheet.getRow(startingRow + 3).getCell(9)).trim();
		}

		// --------------------------------------------------------------------------------------------------------
		// Read Nationality parameters from an excel file
		private static void populateNationalityParameters(Sheet sheet, int startingRow, TestParameters testParameters) {
			// Make sure we are reading the right block of data
			String firstRowCellText = sheet.getRow(startingRow).getCell(0).getRichStringCellValue().getString();
			Assert.assertTrue(firstRowCellText.compareTo("#Nationality") == 0);

			startingRow = startingRow + 1;
			// Get parameters
			DataFormatter formatter = new DataFormatter(); // creating formatter
															// using the default
															// locale
			testParameters.NationalityMYPrimary = formatter.formatCellValue(sheet.getRow(startingRow + 2).getCell(1)).trim();
			testParameters.NationalityPSPrimary = formatter.formatCellValue(sheet.getRow(startingRow + 2).getCell(2)).trim();
			testParameters.NationalityMYPrimaryUpdated = formatter
					.formatCellValue(sheet.getRow(startingRow + 2).getCell(3)).trim();
			testParameters.NationalityPSPrimaryUpdated = formatter
					.formatCellValue(sheet.getRow(startingRow + 2).getCell(4)).trim();
			testParameters.NationalityEmplidPrimary = formatter.formatCellValue(sheet.getRow(startingRow + 2).getCell(5))
					.trim();
			testParameters.NationalityEmplidPrimaryUpdate = formatter
					.formatCellValue(sheet.getRow(startingRow + 2).getCell(6)).trim();
			testParameters.NationalityEmplidPrimaryDelete = formatter
					.formatCellValue(sheet.getRow(startingRow + 2).getCell(7)).trim();

			testParameters.NationalityMYOther = formatter.formatCellValue(sheet.getRow(startingRow + 3).getCell(1)).trim();
			testParameters.NationalityPSOther = formatter.formatCellValue(sheet.getRow(startingRow + 3).getCell(2)).trim();
			testParameters.NationalityMYOtherUpdated = formatter.formatCellValue(sheet.getRow(startingRow + 3).getCell(3)).trim();
			testParameters.NationalityPSOtherUpdated = formatter.formatCellValue(sheet.getRow(startingRow + 3).getCell(4)).trim();
			testParameters.NationalityEmplidOther = formatter.formatCellValue(sheet.getRow(startingRow + 3).getCell(5))
					.trim();
			testParameters.NationalityEmplidOtherUpdate = formatter
					.formatCellValue(sheet.getRow(startingRow + 3).getCell(6)).trim();
			testParameters.NationalityEmplidOtherDelete = formatter
					.formatCellValue(sheet.getRow(startingRow + 3).getCell(7)).trim();
		}

		// --------------------------------------------------------------------------------------------------------
		// Read Language parameters from an excel file
		private static void populateLanguageParameters(Sheet sheet, int startingRow, TestParameters testParameters) {
			// Make sure we are reading the right block of data
			String firstRowCellText = sheet.getRow(startingRow).getCell(0).getRichStringCellValue().getString();
			Assert.assertTrue(firstRowCellText.compareTo("#Language") == 0);

			startingRow = startingRow + 1;
			// Get parameters
			DataFormatter formatter = new DataFormatter(); // creating formatter
															// using the default
															// locale

			// Native
			testParameters.LanguageMYNative = formatter.formatCellValue(sheet.getRow(startingRow + 2).getCell(1)).trim();
			testParameters.LanguagePSNative = formatter.formatCellValue(sheet.getRow(startingRow + 2).getCell(2)).trim();
			testParameters.LanguageMYNativeUpdated = formatter.formatCellValue(sheet.getRow(startingRow + 2).getCell(3)).trim();
			testParameters.LanguagePSNativeUpdated = formatter.formatCellValue(sheet.getRow(startingRow + 2).getCell(4)).trim();
			testParameters.LanguageEmplidNative = formatter.formatCellValue(sheet.getRow(startingRow + 2).getCell(5))
					.trim();
			testParameters.LanguageEmplidNativeUpdate = formatter.formatCellValue(sheet.getRow(startingRow + 2).getCell(6))
					.trim();
			testParameters.LanguageEmplidNativeDelete = formatter.formatCellValue(sheet.getRow(startingRow + 2).getCell(7))
					.trim();

			// Validated
			testParameters.LanguageMYValidated = formatter.formatCellValue(sheet.getRow(startingRow + 3).getCell(1)).trim();
			testParameters.LanguagePSValidated = formatter.formatCellValue(sheet.getRow(startingRow + 3).getCell(2)).trim();
			testParameters.LanguageMYValidatedUpdated = formatter.formatCellValue(sheet.getRow(startingRow + 3).getCell(3)).trim();
			testParameters.LanguagePSValidatedUpdated = formatter.formatCellValue(sheet.getRow(startingRow + 3).getCell(4)).trim();
			testParameters.LanguageEmplidValidated = formatter.formatCellValue(sheet.getRow(startingRow + 3).getCell(5))
					.trim();
			testParameters.LanguageEmplidValidatedUpdate = formatter
					.formatCellValue(sheet.getRow(startingRow + 3).getCell(6)).trim();
			testParameters.LanguageEmplidValidatedDelete = formatter
					.formatCellValue(sheet.getRow(startingRow + 3).getCell(7)).trim();

			// Other
			testParameters.LanguageMYOther = formatter.formatCellValue(sheet.getRow(startingRow + 4).getCell(1)).trim();
			testParameters.LanguagePSOther = formatter.formatCellValue(sheet.getRow(startingRow + 4).getCell(2)).trim();
			testParameters.LanguageMYOtherUpdated = formatter.formatCellValue(sheet.getRow(startingRow + 4).getCell(3)).trim();
			testParameters.LanguagePSOtherUpdated = formatter.formatCellValue(sheet.getRow(startingRow + 4).getCell(4)).trim();
			testParameters.LanguageEmplidOther = formatter.formatCellValue(sheet.getRow(startingRow + 4).getCell(5)).trim();
			testParameters.LanguageEmplidOtherUpdate = formatter.formatCellValue(sheet.getRow(startingRow + 4).getCell(6))
					.trim();
			testParameters.LanguageEmplidOtherDelete = formatter.formatCellValue(sheet.getRow(startingRow + 4).getCell(7))
					.trim();
		}

		// --------------------------------------------------------------------------------------------------------
		// Read Work Experience parameters from an excel file
		private static void populateWorkExperienceParameters(Sheet sheet, int startingRow, TestParameters testParameters) {
			// Make sure we are reading the right block of data
			String firstRowCellText = sheet.getRow(startingRow).getCell(0).getRichStringCellValue().getString();
			Assert.assertTrue(firstRowCellText.compareTo("#WorkExperience") == 0);

			// startingRow = startingRow + 1;
			// Get parameters
			DataFormatter formatter = new DataFormatter(); // creating formatter
															// using the default
															// locale
			// MainJob -Create
			testParameters.MJEmployerID = formatter.formatCellValue(sheet.getRow(startingRow + 2).getCell(1)).trim();
			testParameters.MJEmployerDesc = formatter.formatCellValue(sheet.getRow(startingRow + 2).getCell(2)).trim();
			testParameters.MJStartDate = formatter.formatCellValue(sheet.getRow(startingRow + 2).getCell(3)).trim();
			testParameters.MJEndDate = formatter.formatCellValue(sheet.getRow(startingRow + 2).getCell(4)).trim();
			testParameters.MJJobTitle = formatter.formatCellValue(sheet.getRow(startingRow + 2).getCell(5)).trim();
			testParameters.MJEmplid = formatter.formatCellValue(sheet.getRow(startingRow + 2).getCell(6)).trim();

			// NonMainJob - Create
			testParameters.NonMJEmployerID = formatter.formatCellValue(sheet.getRow(startingRow + 3).getCell(1)).trim();
			testParameters.NonMJEmployerDesc = formatter.formatCellValue(sheet.getRow(startingRow + 3).getCell(2)).trim();
			testParameters.NonMJStartDate = formatter.formatCellValue(sheet.getRow(startingRow + 3).getCell(3)).trim();
			testParameters.NonMJEndDate = formatter.formatCellValue(sheet.getRow(startingRow + 3).getCell(4)).trim();
			testParameters.NonMJJobTitle = formatter.formatCellValue(sheet.getRow(startingRow + 3).getCell(5)).trim();
			testParameters.NonMJEmplid = formatter.formatCellValue(sheet.getRow(startingRow + 3).getCell(6)).trim();

			// Main job - Update
			testParameters.MJEmployerIDUpdate = formatter.formatCellValue(sheet.getRow(startingRow + 4).getCell(1)).trim();
			testParameters.MJEmployerDescUpdate = formatter.formatCellValue(sheet.getRow(startingRow + 4).getCell(2))
					.trim();
			testParameters.MJStartDateUpdate = formatter.formatCellValue(sheet.getRow(startingRow + 4).getCell(3)).trim();
			testParameters.MJEndDateUpdate = formatter.formatCellValue(sheet.getRow(startingRow + 4).getCell(4)).trim();
			testParameters.MJJobTitleUpdate = formatter.formatCellValue(sheet.getRow(startingRow + 4).getCell(5)).trim();
			testParameters.MJEmplidUpdate = formatter.formatCellValue(sheet.getRow(startingRow + 4).getCell(6)).trim();

			// Non Main job - Update
			testParameters.NonMJEmployerIDUpdate = formatter.formatCellValue(sheet.getRow(startingRow + 5).getCell(1))
					.trim();
			testParameters.NonMJEmployerDescUpdate = formatter.formatCellValue(sheet.getRow(startingRow + 5).getCell(2))
					.trim();
			testParameters.NonMJStartDateUpdate = formatter.formatCellValue(sheet.getRow(startingRow + 5).getCell(3)).trim();
			testParameters.NonMJEndDateUpdate = formatter.formatCellValue(sheet.getRow(startingRow + 5).getCell(4)).trim();
			testParameters.NonMJJobTitleUpdate = formatter.formatCellValue(sheet.getRow(startingRow + 5).getCell(5)).trim();
			testParameters.NonMJEmplidUpdate = formatter.formatCellValue(sheet.getRow(startingRow + 5).getCell(6)).trim();

			// Delete
			testParameters.MJEmplidDelete = formatter.formatCellValue(sheet.getRow(startingRow + 7).getCell(1)).trim();
			testParameters.NonMJEmplidDelete = formatter.formatCellValue(sheet.getRow(startingRow + 8).getCell(1)).trim();
		}
		
		// --------------------------------------------------------------------------------------------------------
				// Read Phone parameters from an excel file
				private static void populateMYPhoneParameters(Sheet sheet, int startingRow, TestParameters testParameters) {
					// Make sure we are reading the right block of data
					String firstRowCellText = sheet.getRow(startingRow).getCell(0).getRichStringCellValue().getString();
					Assert.assertTrue(firstRowCellText.compareTo("#Phone") == 0);

					startingRow = startingRow + 1;
					// Get parameters
					DataFormatter formatter = new DataFormatter(); // creating formatter
																	// using the default
																	// locale
				
					testParameters.MYHomePhoneType = formatter.formatCellValue(sheet.getRow(startingRow + 1).getCell(0)).trim();
					testParameters.MYHomePhoneNum = formatter.formatCellValue(sheet.getRow(startingRow + 1).getCell(1)).trim();
					testParameters.MYHomePhoneCountry = formatter.formatCellValue(sheet.getRow(startingRow + 1).getCell(2)).trim();
					testParameters.MYHomePhoneNumUpdated = formatter.formatCellValue(sheet.getRow(startingRow + 1).getCell(3)).trim();
					testParameters.MYHomePhoneCountryUpdated = formatter.formatCellValue(sheet.getRow(startingRow + 1).getCell(4)).trim();
					testParameters.MYHomeEmplidCreate = formatter.formatCellValue(sheet.getRow(startingRow + 1).getCell(5)).trim();
					testParameters.MYHomeEmplidUpdate = formatter.formatCellValue(sheet.getRow(startingRow + 1).getCell(6)).trim();
					testParameters.MYHomeEmplidDelete = formatter.formatCellValue(sheet.getRow(startingRow + 1).getCell(7)).trim();
					
					testParameters.MYBusinessPhoneType = formatter.formatCellValue(sheet.getRow(startingRow + 2).getCell(0)).trim();
					testParameters.MYBusinessPhoneNum = formatter.formatCellValue(sheet.getRow(startingRow + 2).getCell(1)).trim();
					testParameters.MYBusinessPhoneCountry = formatter.formatCellValue(sheet.getRow(startingRow + 2).getCell(2)).trim();
					testParameters.MYBusinessPhoneNumUpdated = formatter.formatCellValue(sheet.getRow(startingRow + 2).getCell(3)).trim();
					testParameters.MYBusinessPhoneCountryUpdated = formatter.formatCellValue(sheet.getRow(startingRow + 2).getCell(4)).trim();
					testParameters.MYBusinessEmplidCreate = formatter.formatCellValue(sheet.getRow(startingRow + 2).getCell(5)).trim();
					testParameters.MYBusinessEmplidUpdate = formatter.formatCellValue(sheet.getRow(startingRow + 2).getCell(6)).trim();
					testParameters.MYBusinessEmplidDelete = formatter.formatCellValue(sheet.getRow(startingRow + 2).getCell(7)).trim();

					testParameters.MYMobilePhoneType = formatter.formatCellValue(sheet.getRow(startingRow + 3).getCell(0)).trim();
					testParameters.MYMobilePhoneNum = formatter.formatCellValue(sheet.getRow(startingRow + 3).getCell(1)).trim();
					testParameters.MYMobilePhoneCountry = formatter.formatCellValue(sheet.getRow(startingRow + 3).getCell(2)).trim();
					testParameters.MYMobilePhoneNumUpdated = formatter.formatCellValue(sheet.getRow(startingRow + 3).getCell(3)).trim();
					testParameters.MYMobilePhoneCountryUpdated = formatter.formatCellValue(sheet.getRow(startingRow + 3).getCell(4)).trim();
					testParameters.MYMobileEmplidCreate = formatter.formatCellValue(sheet.getRow(startingRow + 3).getCell(5)).trim();
					testParameters.MYMobileEmplidUpdate = formatter.formatCellValue(sheet.getRow(startingRow + 3).getCell(6)).trim();
					testParameters.MYMobileEmplidDelete = formatter.formatCellValue(sheet.getRow(startingRow + 3).getCell(7)).trim();
				}
	
}

