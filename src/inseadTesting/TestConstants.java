package inseadTesting;

// Constants to use for the testing session. Offers an easy way of changing the 
// testing environment. e.g. Changing the server to be tested on
public class TestConstants {
	
	// The expected title when we get to the login page of MyInsead
	public static String mMyInseadLoginPageTitle = "Sign in to the INSEAD Network (ADFS-01)";
	
	// The expected title after a successful login
	public static String mMyInseadSuccessfulLoginTitle = "INSEAD - MyINSEAD";
	
	// The expected title after a successful logout
	public static String mMyInseadSuccessfulLogoutTitle = "Sign-Out Page";
	
	// The expected title when we get to the login page of Peoplesoft
	public static String mPeoplesoftLoginPageTitle = "Oracle PeopleSoft Sign-in";
	
	// The expected title when we get to the login page of MailChimp
	public static String mMailChimpLoginTitle = "Login | MailChimp - email marketing made easy";
	
	// The expected title after a successful login to MailChimp
	public static String mMailChimpSuccessfulLoginTitle = "MailChimp Dashboard | Insead";
	
	// The excel file used to load test data
	public static String mTestDataExcelFile = "TestData.xlsx";
	
	// LinkedIn expected title when we get to the login page
	public static String linkedInLoginPageTitle = "Authorize | LinkedIn";
	
	// INT Banner
		public static String bannerintexp = "background-color: rgb(0, 73, 144); background-image: none;";

		// UAT Banner
		public static String banneruatexp = "background-color: rgb(211, 34, 42); background-image: none;";
}
