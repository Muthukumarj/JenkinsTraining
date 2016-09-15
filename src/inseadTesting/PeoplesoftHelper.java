package inseadTesting;

import java.util.List;
import java.util.concurrent.TimeUnit;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Assert;

// --------------------------------------------------------------------------------------------------------
// Class to help with logging in to a MyInsead account
public class PeoplesoftHelper {

	// --------------------------------------------------------------------------------------------------------
	// Log in to MyInsead with a provided account
	public static void doLogin(WebDriver driver, TestParameters testParameters) 
	{
		// Navigate to the login page
		driver.get(testParameters.mURLPeoplesoft);
		WebDriverWait wait = new WebDriverWait(driver, 30);
		wait.until(ExpectedConditions.titleIs(TestConstants.mPeoplesoftLoginPageTitle));
		Assert.assertEquals(driver.getTitle(), TestConstants.mPeoplesoftLoginPageTitle);

		// Populate the username field
		WebElement usernameTextbox = driver.findElement(By.id("userid"));
		usernameTextbox.clear();
		usernameTextbox.sendKeys(testParameters.mPeoplesoftLogin);

		// Populate the password field
		WebElement passwordTextbox = driver.findElement(By.id("pwd"));
		passwordTextbox.clear();
		passwordTextbox.sendKeys(testParameters.mPeoplesoftPassword);

		// Click the sign in button
		WebElement signinButton = driver.findElement(By.name("Submit"));
		signinButton.click();

		// Check that the title is correct
		wait.until(ExpectedConditions.titleIs("Employee-facing registry content"));
	}

	// --------------------------------------------------------------------------------------------------------
	// Synchronizes data from MyINSEAD to Peoplesoft
	public static void syncDataMyINSEADToPeoplesoft(WebDriver driver, MyInseadUser myInseadUser) throws InterruptedException
	{
		executeSynchronizationProcess(driver, myInseadUser, "F1_MY_IN");
	}

	// --------------------------------------------------------------------------------------------------------
	// Synchronizes data from Peoplesoft to MyINSEAD
	public static void syncDataPeoplesoftToMyINSEAD(WebDriver driver, MyInseadUser myInseadUser) throws InterruptedException
	{
		executeSynchronizationProcess(driver, myInseadUser, "F1_MY_OUT");
	}

	// --------------------------------------------------------------------------------------------------------
	// Executes pending processes for a user so that MyINSEAD and PeopleSoft are in sync
	// syncProcessName: F1_MY_IN means MyINSEAD data is copied to Peoplesoft
	//					F1_MY_OUT means Peoplesoft data is copied to MyINSEAD
	private static void executeSynchronizationProcess(WebDriver driver, MyInseadUser myInseadUser, String syncProcessName) throws InterruptedException 
	{
		// Make sure syncProcessName is what we expect
		Assert.assertTrue(syncProcessName.compareTo("F1_MY_IN") == 0 || syncProcessName.compareTo("F1_MY_OUT")  == 0, "syncProcessName must be F1_MY_IN or F1_MY_OUT, was " + syncProcessName);


		// Assumes that we are logged in to PeopleSoft
		WebDriverWait wait = new WebDriverWait(driver, 30);
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("pthnavbca_PORTAL_ROOT_OBJECT")));
		driver.findElement(By.id("pthnavbca_PORTAL_ROOT_OBJECT")).click();
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("fldra_F1_MYINSEAD_SYNC")));
		driver.findElement(By.id("fldra_F1_MYINSEAD_SYNC")).click();
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.linkText("MYINSEAD Integration")));
		driver.findElement(By.linkText("MYINSEAD Integration")).click();

		driver.switchTo().frame("TargetContent");
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("PRCSRUNCNTL_RUN_CNTL_ID")));
		driver.findElement(By.id("PRCSRUNCNTL_RUN_CNTL_ID")).sendKeys("MYINSEAD_TESTAUTOMATION");
		driver.findElement(By.id("#ICSearch")).click();
		wait.until(ExpectedConditions.invisibilityOfElementLocated(By.id("processing")));
		
		boolean processnamenotxisting = isElementPresent(By.id("ICTAB_1"), driver);
		if (processnamenotxisting==true)
		{
			// Create a new process item
			wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ICTAB_1")));
			driver.findElement(By.id("ICTAB_1")).click();
			wait.until(ExpectedConditions.invisibilityOfElementLocated(By.id("#ICMatchCase")));
			wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("PRCSRUNCNTL_RUN_CNTL_ID")));
			driver.findElement(By.id("PRCSRUNCNTL_RUN_CNTL_ID")).clear();
			String newProcessName = "MYINSEAD_TESTAUTOMATION";
			driver.findElement(By.id("PRCSRUNCNTL_RUN_CNTL_ID")).sendKeys(newProcessName);
			driver.findElement(By.id("#ICSearch")).click();
			wait.until(ExpectedConditions.invisibilityOfElementLocated(By.id("WAIT.win0")));
			List<WebElement> existingProcessControlIDs = driver.findElements(By.id("SEARCH_RESULT1"));
			if (!existingProcessControlIDs.isEmpty())
			{
				// We already have a process control ID by this name. Click it and proceed.
				existingProcessControlIDs.get(0).click();
			}
		}
		else{
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("F1_MY_SYNC_RUN_DEBUG")));
		}
		
		// Run incoming and outgoing processes for the given id
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("PRCSRQSTDLG_WRK_LOADPRCSRQSTDLGPB")));
		driver.findElement(By.id("PRCSRQSTDLG_WRK_LOADPRCSRQSTDLGPB")).click();
		if (0 == syncProcessName.compareTo("F1_MY_IN"))
		{
			wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("PRCSRQSTDLG_WRK_SELECT_FLAG$0")));
			driver.findElement(By.id("PRCSRQSTDLG_WRK_SELECT_FLAG$0")).click();
		}
		else if (0 == syncProcessName.compareTo("F1_MY_OUT"))
		{
			wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("PRCSRQSTDLG_WRK_SELECT_FLAG$1")));
			driver.findElement(By.id("PRCSRQSTDLG_WRK_SELECT_FLAG$1")).click();
		}
		driver.findElement(By.id("#ICSave")).click();
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("PRCSRQSTDLG_WRK_LOADPRCSMONITORPB")));
		driver.findElement(By.id("PRCSRQSTDLG_WRK_LOADPRCSMONITORPB")).click();

		// At this point we have asked Peoplesoft to run the processes. Keep monitoring (and clicking "Refresh")
		// until we see the "Success" message in the table. We can't just wait for the text to the "Success" as
		// we have to click refresh to get new messages.
		boolean bExecutionSuccess = false;
		int secondsToWaitPerRefresh = 20, secondsToWaitForProcessesToExecute = 1000;
		int numRefreshes = 0, maxNumRefreshes = secondsToWaitForProcessesToExecute / secondsToWaitPerRefresh; 

		String processId = "";

		driver.manage().timeouts().implicitlyWait(secondsToWaitPerRefresh, TimeUnit.SECONDS);
		Thread.sleep(2000);
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("JOBDETAIL_BTN$span$0")));
		if (0 == syncProcessName.compareTo(driver.findElement(By.id("JOBDETAIL_BTN$span$0")).getText())) {
			processId = driver.findElement(By.id("PMN_PRCSLIST_PRCSINSTANCE$0")).getText();
String index = "";
			while (!bExecutionSuccess && ++numRefreshes < maxNumRefreshes)
			{
				// Keep in this loop until we get messages stating that the processes executed successfully.
				// One exception is: If we have refreshes a lot of times and still don't get the message something
				// might be wrong and we have to exit the test.

				// Wait for 5 seconds
				driver.manage().timeouts().implicitlyWait(secondsToWaitPerRefresh, TimeUnit.SECONDS);

				// Click refresh
				driver.findElement(By.id("REFRESH_BTN")).click();
				wait.until(ExpectedConditions.invisibilityOfElementLocated(By.id("WAIT.win0")));

				WebElement process = driver.findElement(By.xpath("//span[contains(text(),'" + processId + "')]"));
				String id = process.getAttribute("id");
				String val = id.substring(id.lastIndexOf('$')+1);
				if (0 == syncProcessName.compareTo("F1_MY_IN")) {
					Assert.assertEquals(driver.findElement(By.id("JOBDETAIL_BTN$span$"+val)).getText(), "F1_MY_IN");
					// "IN" process is a success if the status description in the table says "Success"
					bExecutionSuccess = "Success".compareTo(driver.findElement(By.id("PMN_PRCSLIST_RUNSTATUSDESCR$"+val)).getText()) == 0;
				}
				else if (0 == syncProcessName.compareTo("F1_MY_OUT")) {
					Assert.assertEquals(driver.findElement(By.id("JOBDETAIL_BTN$span$"+val)).getText(), "F1_MY_OUT");
					// "OUT" process is a success if the status description in the table says "Success"
					bExecutionSuccess = "Success".compareTo(driver.findElement(By.id("PMN_PRCSLIST_RUNSTATUSDESCR$"+val)).getText()) == 0;
				}
				else {
					Assert.assertTrue(false, "Unexpected condition when synchronizing data");
				}
				index = val;
			}
			Assert.assertEquals(driver.findElement(By.id("PMN_PRCSLIST_RUNSTATUSDESCR$" + index)).getText(), "Success");
		} else {
			Assert.assertTrue(false, "Unexpected condition when synchronizing data");
		}

	}
	
	//-------------------------------------------------------------------------------------------------------------
	// Validate Sync monitor content
	public static void SyncMonitorDataValidation(WebDriver driver, String emplid, String contenttype, String content, TestParameters testParameters, String direction)
	throws InterruptedException {
	driver.get(testParameters.mURLMyInsead+testParameters.msyncMonitor);
	Assert.assertEquals(driver.getTitle(), "MyINSEAD - Sync Monitor");
	
	int lastindex = -1;
	boolean emplidfound = false;
	boolean dataprocessed = false;
	for (int i = 1; i <= 25; i++) {
		System.out.println("\nIteration no.: " + i);
		WebElement key = driver.findElement(By.xpath("(//tbody/tr[" + i + "]/td)[5]"));
		String keydata = key.getText();
		System.out.println(keydata);
		boolean result = keydata.contains("Key: " + emplid);
		String actualDirection = driver.findElement(By.xpath("//*[@id='tblResult']/tbody/tr["+i+"]/td[3]")).getText().trim();
		Thread.sleep(5000);
		
		if (result == true && actualDirection.equals(direction)) { // Validate if emplid is in the key
			emplidfound = true;
			String status = driver.findElement(By.xpath("(//tbody/tr[" + i + "]/td)[2]")).getText();
			System.out.println(status);
			boolean statusresult = status.equals("PROCESSED");
			String dataID = driver.findElement(By.xpath("(//tbody/tr["+i+"]/td)[1]")).getText();//TODO get the data id
			System.out.println(dataID);
		
			//Validation if Data is processed
		if (statusresult == true) {// Validate if processed
			System.out.println("Data PROCESSED");
			dataprocessed = true;
			lastindex = i;
			break;
		} else {
			System.out.println("Data Not Yet Processed");
			int count = 20;
			String dataIDafterrefresh = driver.findElement(By.xpath("(//tbody/tr["+i+"]/td)[1]")).getText();//TODO to remove
			boolean correctdataid = dataID.equals(dataIDafterrefresh);//TODO to remove
			System.out.println(correctdataid);//TODO to remove
			
			while (count != 0 && statusresult == false) {// Loop for
			// processed
				
				if(correctdataid==true){
			System.out.println("Status check: " + status);
			System.out.println(count);
			driver.navigate().refresh();
			Thread.sleep(15000);
			status = driver.findElement(By.xpath("(//tbody/tr[" + i + "]/td)[2]")).getText();
			statusresult = status.equals("PROCESSED");
			count--;
			}else
			{
				//TODO GO BACK TO CHECK EMPLID
			}
			}		
		if(statusresult==true)
		{
			lastindex = i;
			dataprocessed = true;
			break;
		}
		}
		break;
		} else {
		System.out.println("Checking next data");
		}
	}
	
		if(!emplidfound)
		{
			System.out.println("No data in Sync Monitor");
			Assert.assertTrue(emplidfound);
		}
		if(dataprocessed)
		{
			Thread.sleep(3000);
			WebElement button = driver.findElement(By.xpath("(//tbody/tr[" + lastindex + "]/td)[1]/button"));
			
			// Get data content
			String dataBlock = button.getAttribute("data-content").replaceAll("\\s+", "");

			// Validate that content was correct
			boolean contenttypefound = dataBlock.contains(contenttype);
			boolean contentfound = dataBlock.contains(content);
		
			Assert.assertEquals(true, contenttypefound);
			Assert.assertEquals(true, contentfound);
		}else{
			System.out.println("Data was not processed in sync monitor");
			Assert.assertTrue(dataprocessed);
		}
		
		}

public static boolean isElementPresent(By by, WebDriver driver) {
	try {
		driver.findElement(by);
		return true;
	} catch (org.openqa.selenium.NoSuchElementException e) {
		return false;
	}
}
	
}
