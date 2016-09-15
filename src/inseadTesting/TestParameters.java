package inseadTesting;

// --------------------------------------------------------------------------------------------------------
// The test parameters (e.g. the URL for MyINSEAD) to be used for testing
public class TestParameters {

	// The MyInsead server URL
	public String mURLMyInsead = "";
	
	// The MyInsead server URL
		public String msyncMonitor =  "/sync-monitor/data";

	// Username for a user with global admin rights in MyINSEAD
	public String mMyInseadGlobalAdminLogin = "";

	// Password for a user with global admin rights in MyINSEAD
	public String mMyInseadGlobalAdminPassword = "";

	// The Peoplesoft server URL
	public String mURLPeoplesoft = "";

	// Login for the Peoplesoft account
	public String mPeoplesoftLogin = "";

	// Password for the Peoplesoft account
	public String mPeoplesoftPassword = "";

	// The MailChimp URL
	public String mURLMailChimp = "https://login.mailchimp.com/";

	// Login for MailChimp
	public String mMailChimpLogin = "";

	// Password for MailChimp
	public String mMailChimpPassword = "";

	// Phone number for the FUNC326001_2_Update_Business_Phone_in_MyINSEAD test
	public String mFUNC326001_2_BusinessPhoneCountryCode = "";
	public String mFUNC326001_2_BusinessPhone = "";

	// Parameters for the TC_DirectorySearch_ByFirstName test
	public String mTC_DirectorySearch_ByFirstName_Name = "";
	public String mTC_DirectorySearch_ByFirstName_Results = "";

	// Employee IDs for users in the "Roles" test cases
	public String mRolesAlumni = "";
	public String mRolesStudent = "";
	public String mRolesStaffFaculty = "";
	public String mRolesEDP = "";
	public String mRolesMBAAdmin = "";
	public String mRolesGlobalAdmin = "";
	public String mRolesNAAAdmin = "";
	public String mRolesAffiliate = "";
	public String AlumniFirstName = "";
	public String AlumniLastName = "";
	public String StudentS52 = "";
	public String StudentFirstName = "";
	public String StudentLastName  = "";
	
	// Student Logins for EMBA/MBA Promotion settings => Drowdown of students and Login As button
	public String mRolesEMBAstudentByName = "";
	
	// affiliate roles
	public String MIS50_STRY0010302_AffiliateViewMember="";
	public String MIS50_STRY0010302_AffiliateViewNonMember="";
	public String MIS50_STRY0010302_AlumniMemberNAAadmin="";
	public String MIS50_STRY0010302_AlumniNonMemberNAAadmin="";
	public String MIS50_STRY0010302_AffiliateMemberNAAadmin="";
	public String MIS50_STRY0010302_AffiliateNonMemberNAAadmin="";
	public String MIS50_STRY0010302_AffiliateNonNAAadmin="";
	public String MIS50_STRY0010302_AffiliateOtherNAAadmin="";
	public String MIS50_STRY0010302_AffiliateOtherNonNAAadmin="";
	
	// Parameters for the TC_MailChimp_NonMember test
	public String mTC_MailChimp_NonMember_EMPLID = "";
	public String mTC_MailChimp_NonMember_List = "";

	// The MyInsead URL
	public String myInseadURL = "https://my-int.insead.edu";

	// The MyInsead URL
	public String myInseadURLdashboard = "https://my-int.insead.edu/dashboard/";

	// Employee IDs for users in the "EMBAStudent" test cases
	public String EMBAStudent1 = "";
	public String EMBAStudent2 = "";

	// Employee IDs for users without preferred flags
	public String woPreferredPhone;
	public String woPreferredEmail;
	public String woPreferredAddress;
	public String wHomePreferredAddress;
	
	// Employee IDs for users in the "EMBAStudent" test cases
			public String eMBAStudent1 = "";
			public String embaStudent1Upn ="";
			public String eMBAStudent2 = "";
			public String embaStudent2Upn ="";
			public String eMBAStudent3 = "";
			public String embaStudent3Upn ="";
			public String eMBAStudent4 = "";
			public String embaStudent4Upn ="";
			public String eMBAStudent5 = "";
			public String embaStudent5Upn ="";
			
			public String embaStudentSGP="";
			public String embaStudentSGPUPN="";
			
			// EMBA Student2 course data 
			public String classNbrforEmbaStu2Cor = "";
			public String corCourseNamewithSectionStu2 = "";
			
			public String classNbrforEmbaStu2Ele = "";
			public String eleCourseNamewithSectionStu2 = "";
			
			public String classNbrforEmbaStu2Kmc = "";
			public String kmcCourseNamewithSectionStu2 = "";
			
			public String classNbrforEmbaStu2Pro = "";
			public String proCourseNamewithSectionStu2 = "";
			
			// EMBA Student3 course data 
			public String classNbrforEmbaStu3Cor = "";
			public String corCourseNamewithSectionStu3 = "";
					
			public String classNbrforEmbaStu3Ele = "";
			public String eleCourseNamewithSectionStu3 = "";
					
			public String classNbrforEmbaStu3Kmc = "";
			public String kmcCourseNamewithSectionStu3 = "";
					
			public String classNbrforEmbaStu3Pro = "";
			public String proCourseNamewithSectionStu3 = "";
			
			// EMBA Student4 course data 
			public String classNbrforEmbaStu4Cor = "";
			public String corCourseNamewithSectionStu4 = "";
					
			public String classNbrforEmbaStu4Ele = "";
			public String eleCourseNamewithSectionStu4 = "";
						
			public String classNbrforEmbaStu4Kmc = "";
			public String kmcCourseNamewithSectionStu4 = "";
					
			public String classNbrforEmbaStu4Pro = "";
			public String proCourseNamewithSectionStu4 = "";
			
			// test data related to MIS50_STRY0010675_MyInseadJobTitleChangefrUpperCase
			
			// User1(Alumni or Student)- Job title all in upper case
			public String student_or_alumni_user1_emplid;
			
			// User2(Alumni or Student)- Job title all in lower case
			public String student_or_alumni_user2_emplid;
			
			// User3(Alumni or Student)- Job title all in upper case (LinkedIn -tests)
			public String student_or_alumni_user3_emplid;
			
			// User4(Alumni or Student)- Job title all in lower case (LinkedIn -tests)
			public String student_or_alumni_user4_emplid;
			
			// User5(Alumni or Student)- Job title all in upper case (Use for PeopleSoft tests)
			public String student_or_alumni_user5_emplid;
						
			// User6(Alumni or Student)- Job title all in upper case (Use for PeopleSoft tests)
			public String student_or_alumni_user6_emplid;			
			
			// LinkedIn username
			public String linkedInUserName;
			
			// LinkedIn password 
			public String linkedInPassword;	
			
			// test country for jobForm via LinkedIn 
			public String countryForJobForm;
			
			// test industry for jobForm via LinkedIn
			public String industryForJobForm;
			
			// test company for jobForm via LinkedIn 
			public String companyForJobForm;
			
			// Term Campus Deletion EMPLID
			public String termstudent = "";
			
			// test parameters for STRY0010723 (Activation_Deactivation_MembershipReminder) story 
			// Alumni who's membership expires within 30 days
			public String alumni_user1 ="";
			
			// Alumni who's membership expires within 7 days
			public String alumni_user2 ="";
						
			// Alumni who's membership expires within 0 days
			public String alumni_user3 ="";
			
			// NAA Admin 1 with membership reminder enabled 
			public String naa_admin_for_user1_2_3_10_11_12 ="";
			
			// Alumni who's membership(IAA) expires within 30 days
			public String alumni_user4 ="";
			
			// Alumni who's membership(IAA) expires within 7 days
			public String alumni_user5 ="";
			
			// Alumni who's membership(IAA) expires within 0 days
			public String alumni_user6 ="";
			
			// IAA Admin 1 with membership reminder enabled 
			public String naa_admin_for_iaa ="";
			
			// Alumni who has multiple memberships and one of his membership expires within 30 days
			public String alumni_user7 ="";
			
			// Alumni who has multiple memberships and one of his membership expires within 7 days
			public String alumni_user8 ="";
			
			// Alumni who has multiple memberships and one of his membership expires within 0 days
			public String alumni_user9 ="";
			
			// NAA Admin with membership reminder enabled 
			public String naa_admin_for_user7_8_9_16_17_18 ="";
			
			// Membership reminder disabled from user_setting for following users
			
			// Alumni who has one membership and his membership expires within 30 days - Reminder User_setting disabled  
			public String alumni_user10 ="";
			
			// Alumni who has one membership and his membership expires within 7 days - Reminder User_setting disabled  
			public String alumni_user11 ="";
			
			// Alumni who has one membership and his membership expires within 0 days - Reminder User_setting disabled  
			public String alumni_user12 ="";
			
			// Alumni who's membership(IAA) expires within 30 days - Reminder User_setting disabled  
			public String alumni_user13 ="";

			// Alumni who's membership(IAA) expires within 7 days - Reminder User_setting disabled  
			public String alumni_user14 ="";

			
			// Alumni who's membership(IAA) expires within 0 days - Reminder User_setting disabled  
			public String alumni_user15 ="";
			
			// Alumni who has multiple memberships and one of his membership expires within 30 days- Reminder User_setting disabled
			public String alumni_user16 ="";

			
			// Alumni who has multiple memberships and one of his membership expires within 30 days- Reminder User_setting disabled
			public String alumni_user17 ="";

			
			// Alumni who has multiple memberships and one of his membership expires within 30 days- Reminder User_setting disabled
			public String alumni_user18 ="";
			
			// Alumni that has only 30 days left with the membership and renewed the membership 31 days until the expiration date - Reminder NAA_setting and User_setting enabled
			public String alumni_user28 ="";
			
			// Alumni that has only 7 days left with the membership and renewed the membership 8 days until the expiration date - Reminder NAA_setting and User_setting enabled
			public String alumni_user29 ="";

			// Alumni that is expiring today and renewed the membership 1 day until the expiration date - Reminder NAA_setting and User_setting enabled
			public String alumni_user30 ="";
			
			// Alumni that has only 30 days left with the membership and renewed the membership 31 days until the expiration date - Reminder User_setting disabled
			public String alumni_user31 ="";
			
			// Alumni that has only 7 days left with the membership and renewed the membership 8 days until the expiration date - Reminder User_setting disabled
			public String alumni_user32 ="";

			// Alumni that is expiring today and renewed the membership 1 day until the expiration date - Reminder User_setting disabled
			public String alumni_user33 ="";
			
			// Alumni that has only 30 days left with the membership and renewed the membership 31 days until the expiration date - Reminder NAA_setting disabled
			public String alumni_user34 ="";
			
			// Alumni that has only 7 days left with the membership and renewed the membership 8 days until the expiration date - Reminder NAA_setting disabled
			public String alumni_user35 ="";
			
			// Alumni that is expiring today and renewed the membership 1 day until the expiration date - Reminder NAA_setting disabled
			public String alumni_user36 ="";
			
			
			public String webMailUrl ="https://webmail.insead.edu/";
			
			public String webMailUserName ="";
			
			public String webMailPassword ="";
			
			public String reminderTaskUrl ="";

			// NAA Admin with membership 
			public String naa_admin_for_user_19 ="";

			//Alumni that has only 30 days left with the membership when the setting is NAA setting = Off
			public String alumni_user19 ="";
			
			//Alumni that has only 7 days left with the membership when the setting is NAA setting = Off
			public String alumni_user20 ="";
			
			//Alumni that has only 0 days left with the membership when the setting is NAA setting = Off
			public String alumni_user21 ="";
			
			//User 22(Alumni that has only 30 days left with the IAA membership (NAA setting = OFF))
			public String alumni_user22 ="";
			
			// NAA Admin with membership ( No country )  
			public String naa_admin_for_user_22 ="";
			
			//User 23(Alumni that has only 7 days left with the IAA membership (NAA setting = OFF))
			public String alumni_user23 ="";
			
			//User 24(Alumni that has only 0 days left with the IAA membership (NAA setting = OFF))
			public String alumni_user24 ="";
			
			//User 25(Alumni that has only 30 days left with the multiple membership + (NAA setting = OFF))
			public String alumni_user25 ="";
			
			// NAA Admin
			public String naa_admin_for_user_25 ="";
			
			//User 26(Alumni that has only 7 days left with the multiple membership + (NAA setting = OFF))
			public String alumni_user26 ="";
			
			//User 27(Alumni that has only 0 days left with the multiple membership + (NAA setting = OFF))
			public String alumni_user27 ="";
			
			
			
			// ==========================================================

			// RegressionEmplids
			public String CRUDAlumni;
			public String CRUDStudent;

			// Valid Constituent Types
			public String constituentTypeAlumni;
			public String constituentTypeAlumniValue;
			public String constituentTypeAlumniEmplid;
			public String constituentTypeStudent;
			public String constituentTypeStudentValue;
			public String constituentTypeStudentEmplid;
			public String constituentTypeFaculty;
			public String constituentTypeFacultyValue;
			public String constituentTypeFacultyEmplid;
			public String constituentTypeStaff;
			public String constituentTypeStaffValue;
			public String constituentTypeStaffEmplid;
			public String constituentTypeINSEADClient;
			public String constituentTypeINSEADClientValue;
			public String constituentTypeINSEADClientEmplid;
			public String constituentTypeINSEADContractor;
			public String constituentTypeINSEADContractorValue;
			public String constituentTypeINSEADContractorEmplid;
			public String constituentTypeAffiliate;
			public String constituentTypeAffiliateValue;
			public String constituentTypeAffiliateEmplid;
			public String constituentTypeParticipant;
			public String constituentTypeParticipantValue;
			public String constituentTypeParticipantEmplid;
			public String constituentTypeExchangerStudent;
			public String constituentTypeExchangerStudentValue;
			public String constituentTypeExchangerStudentEmplid;
			public String constituentTypePastParticipant;
			public String constituentTypePastParticipantValue;
			public String constituentTypePastParticipantEmplid;

			// Update
			public String constituentTypeAlumniUpdate;
			public String constituentTypeStudentUpdate;
			public String constituentTypeFacultyUpdate;
			public String constituentTypeStaffUpdate;
			public String constituentTypeINSEADClientUpdate;
			public String constituentTypeINSEADContractorUpdate;
			public String constituentTypeAffiliateUpdate;
			public String constituentTypeParticipantUpdate;
			public String constituentTypeExchangerStudentUpdate;
			public String constituentTypePastParticipantUpdate;

			// Delete
			public String constituentTypeAlumniDelete;
			public String constituentTypeStudentDelete;
			public String constituentTypeFacultyDelete;
			public String constituentTypeStaffDelete;
			public String constituentTypeINSEADClientDelete;
			public String constituentTypeINSEADContractorDelete;
			public String constituentTypeAffiliateDelete;
			public String constituentTypeParticipantDelete;
			public String constituentTypeExchangerStudentDelete;
			public String constituentTypePastParticipantDelete;

			// New Person outputs
			public String AddPersonFirstNameAlumni;
			public String AddPersonLastNameAlumni;
			public String AddPersonEmailAlumni;
			public String AddPersonEmplidAlumni;

			public String AddPersonFirstNameStudent;
			public String AddPersonLastNameStudent;
			public String AddPersonEmailStudent;
			public String AddPersonEmplidStudent;

			public String AddPersonFirstNameStaff;
			public String AddPersonLastNameStaff;
			public String AddPersonEmailStaff;
			public String AddPersonEmplidStaff;

			public String AddPersonFirstNameFaculty;
			public String AddPersonLastNameFaculty;
			public String AddPersonEmailFaculty;
			public String AddPersonEmplidFaculty;

			// Phones
			public String HomeEmplidCreate;
			public String HomeEmplidUpdate;
			public String HomeEmplidDelete;
			public String HomePhoneType;
			public String HomePhoneNum;
			public String HomePhoneExt;
			public String HomePhoneCountry;
			public String HomePhoneNumUpdated;
			public String HomePhoneExtUpdated;
			public String HomePhoneCountryUpdated;

			public String BusinessEmplidCreate;
			public String BusinessEmplidUpdate;
			public String BusinessEmplidDelete;
			public String BusinessPhoneType;
			public String BusinessPhoneNum;
			public String BusinessPhoneExt;
			public String BusinessPhoneCountry;
			public String BusinessPhoneNumUpdated;
			public String BusinessPhoneExtUpdated;
			public String BusinessPhoneCountryUpdated;

			public String MobileEmplidCreate;
			public String MobileEmplidUpdate;
			public String MobileEmplidDelete;
			public String MobilePhoneType;
			public String MobilePhoneNum;
			public String MobilePhoneExt;
			public String MobilePhoneCountry;
			public String MobilePhoneNumUpdated;
			public String MobilePhoneExtUpdated;
			public String MobilePhoneCountryUpdated;

			// Emails
			public String HomeEmailType;
			public String HomeEmailAddress;
			public String HomeEmailAddressUpdated;
			public String HomeEmailEmplidCreate;
			public String HomeEmailEmplidUpdate;
			public String HomeEmailEmplidDelete;

			public String BusinessEmailType;
			public String BusinessEmailAddress;
			public String BusinessEmailAddressUpdated;
			public String BusinessEmailEmplidCreate;
			public String BusinessEmailEmplidUpdate;
			public String BusinessEmailEmplidDelete;

			public String UPNEmailType;
			public String UPNEmailAddress;
			public String UPNEmailAddressUpdated;
			public String UPNEmailEmplidCreate;
			public String UPNEmailEmplidUpdate;
			public String UPNEmailEmplidDelete;

			// Nationality
			public String NationalityMYPrimary;
			public String NationalityPSPrimary;
			public String NationalityMYPrimaryUpdated;
			public String NationalityPSPrimaryUpdated;
			public String NationalityEmplidPrimary;
			public String NationalityEmplidPrimaryUpdate;
			public String NationalityEmplidPrimaryDelete;

			public String NationalityMYOther;
			public String NationalityPSOther;
			public String NationalityMYOtherUpdated;
			public String NationalityPSOtherUpdated;
			public String NationalityEmplidOther;
			public String NationalityEmplidOtherUpdate;
			public String NationalityEmplidOtherDelete;

			// Language
			// Native
			public String LanguageMYNative;
			public String LanguagePSNative;
			public String LanguageMYNativeUpdated;
			public String LanguagePSNativeUpdated;
			public String LanguageEmplidNative;
			public String LanguageEmplidNativeUpdate;
			public String LanguageEmplidNativeDelete;

			// Validated
			public String LanguageMYValidated;
			public String LanguagePSValidated;
			public String LanguageMYValidatedUpdated;
			public String LanguagePSValidatedUpdated;
			public String LanguageEmplidValidated;
			public String LanguageEmplidValidatedUpdate;
			public String LanguageEmplidValidatedDelete;

			// Other
			public String LanguageMYOther;
			public String LanguagePSOther;
			public String LanguageMYOtherUpdated;
			public String LanguagePSOtherUpdated;
			public String LanguageEmplidOther;
			public String LanguageEmplidOtherUpdate;
			public String LanguageEmplidOtherDelete;

			// Work Experience
			// MainJob
			public String MJEmployerID;
			public String MJEmployerDesc;
			public String MJStartDate;
			public String MJEndDate;
			public String MJJobTitle;
			public String MJEmployerIDUpdate;
			public String MJEmployerDescUpdate;
			public String MJStartDateUpdate;
			public String MJEndDateUpdate;
			public String MJJobTitleUpdate;
			public String MJEmplid;
			public String MJEmplidUpdate;
			public String MJEmplidDelete;
			// NonMainJob
			public String NonMJEmployerID;
			public String NonMJEmployerDesc;
			public String NonMJStartDate;
			public String NonMJEndDate;
			public String NonMJJobTitle;
			public String NonMJEmployerIDUpdate;
			public String NonMJEmployerDescUpdate;
			public String NonMJStartDateUpdate;
			public String NonMJEndDateUpdate;
			public String NonMJJobTitleUpdate;
			public String NonMJEmplid;
			public String NonMJEmplidUpdate;
			public String NonMJEmplidDelete;
			

			public String mailchimp_member1 = "";
			
			public String mailchimp_member1_email = "";

			public String mailchimp_member2 = "";
			
			public String mailchimp_member2_email = "";
			
			public String mailchimp_member3 = "";
			
			public String mailchimp_member3_email = "";

			public String mailchimp_member4 = "";
			
			public String mailchimp_member4_email = "";
			
			public String mailchimp_member5 = "";
			
			public String mailchimp_member5_email = "";

			public String mailchimp_member6 = "";
			
			public String mailchimp_member6_email = "";
			
			public String mailchimp_member7 = "";
			
			public String mailchimp_member7_email = "";

			public String mailchimp_member8 = "";
			
			public String mailchimp_member8_email = "";
			
			public String mailchimp_member9 = "";
			
			public String mailchimp_member9_email = "";

			public String mailchimp_member10 = "";
			
			public String mailchimp_member10_email = "";


			public String MYHomeEmplidCreate;
			public String MYHomeEmplidUpdate;
			public String MYHomeEmplidDelete;
			public String MYHomePhoneType;
			public String MYHomePhoneNum;
			public String MYHomePhoneCountry;
			public String MYHomePhoneNumUpdated;
			public String MYHomePhoneCountryUpdated;

			public String MYBusinessEmplidCreate;
			public String MYBusinessEmplidUpdate;
			public String MYBusinessEmplidDelete;
			public String MYBusinessPhoneType;
			public String MYBusinessPhoneNum;
			public String MYBusinessPhoneCountry;
			public String MYBusinessPhoneNumUpdated;
			public String MYBusinessPhoneCountryUpdated;

			public String MYMobileEmplidCreate;
			public String MYMobileEmplidUpdate;
			public String MYMobileEmplidDelete;
			public String MYMobilePhoneType;
			public String MYMobilePhoneNum;
			public String MYMobilePhoneCountry;
			public String MYMobilePhoneNumUpdated;
			public String MYMobilePhoneCountryUpdated;
			
			public String stry0010503_user1= "";
			public String stry0010503_user2 = "";

}
