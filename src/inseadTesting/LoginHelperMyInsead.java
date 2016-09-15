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
public class LoginHelperMyInsead {

	// --------------------------------------------------------------------------------------------------------
	// Log in to MyInsead with a provided account
	public static void loginViaLoginPage(WebDriver driver, TestParameters testParameters) 
	{
		// Navigate to the login page
		driver.get(testParameters.mURLMyInsead);
		  driver.manage().window().maximize();
		WebDriverWait wait = new WebDriverWait(driver, 60);
		wait.until(ExpectedConditions.titleIs(TestConstants.mMyInseadLoginPageTitle));
		Assert.assertEquals(driver.getTitle(), TestConstants.mMyInseadLoginPageTitle);
		
		// Populate the username field
		WebElement usernameTextbox = driver.findElement(By.id("ctl00_ContentPlaceHolder1_UsernameTextBox"));
		usernameTextbox.sendKeys(testParameters.mMyInseadGlobalAdminLogin);
		
		// Populate the password field
		WebElement passwordTextbox = driver.findElement(By.id("ctl00_ContentPlaceHolder1_PasswordTextBox"));
		passwordTextbox.sendKeys(testParameters.mMyInseadGlobalAdminPassword);
		
		// Click the sign in button
		wait.until(ExpectedConditions.elementToBeClickable(By.id("ctl00_ContentPlaceHolder1_SubmitButton")));
		driver.findElement(By.id("ctl00_ContentPlaceHolder1_SubmitButton")).click();
		
		// Check that the title is correct
		wait.until(ExpectedConditions.titleIs(TestConstants.mMyInseadSuccessfulLoginTitle));
		Assert.assertEquals(driver.getTitle(), TestConstants.mMyInseadSuccessfulLoginTitle);
	}

	// --------------------------------------------------------------------------------------------------------
	// Log in to MyInsead as a given user, by using the "Global Admin" link of an administrative login
	public static void loginViaGlobalAdmin(WebDriver driver, MyInseadUser userToLoginAs)
	{
		// Assumption: We are already logged in as a user that is a global admin, and we can see the "Global Admin" link
		// Click on "Global Admin"
		WebElement globalAdminLink = driver.findElement(By.linkText("Global Admin"));
		globalAdminLink.click();
		
		// Wait till the tab loads
		WebDriverWait wait = new WebDriverWait(driver, 60);
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("form_param")));
		
		// Add the employee id of the user to login as
		boolean loginAsUserFound = false;
		
		driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);	// Implicitly wait when finding elements in DOM
		int numTries = 0, maxNumTries = 20;
		while (!loginAsUserFound && ++numTries < maxNumTries)
		{
			WebElement employeeIDField = driver.findElement(By.id("form_param"));
			employeeIDField.clear();
			employeeIDField.sendKeys(userToLoginAs.mEMPLID);
	
			// Click on the result (only one should show up)
			//wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector(".ui-autocomplete.ui-front.ui-menu.ui-widget.ui-widget-content.ui-corner-all")));
			//WebElement autocompleteMenu = driver.findElement(By.cssSelector(".ui-autocomplete.ui-front.ui-menu.ui-widget.ui-widget-content.ui-corner-all"));
			wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector(".ui-autocomplete")));
			WebElement autocompleteMenu = driver.findElement(By.cssSelector(".ui-autocomplete"));
			// Get the first result. Assume this is the right one as we are logging in by employee ID, 
			// so should only have one result.
			WebElement firstResult = autocompleteMenu.findElements(By.tagName("a")).get(0);
			firstResult.click();
			
			// Wait till we get the "Login as User" field
			System.out.println("Trying to get login as user field");
			loginAsUserFound = driver.findElements(By.xpath("//input[@value='Login as User']")).size() > 0;
			System.out.println("Returned " + loginAsUserFound);
			if (!loginAsUserFound) {
				// The "Login as user" button was not found. We are going to navigate to another tab and back to the "User" tab
				driver.findElement(By.id("ui-id-2")).click();
				wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("navGroups")));
				// Navigate back to the "User" tab
				driver.findElement(By.id("ui-id-1")).click();
				wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("form_param")));
			}
		}
		driver.manage().timeouts().implicitlyWait(0, TimeUnit.SECONDS);	// Reset to default
		
		// Click on "login as user"
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//input[@value='Login as User']")));
		WebElement loginButton = driver.findElement(By.xpath("//input[@value='Login as User']"));
		loginButton.click();
		wait.until(ExpectedConditions.titleIs(TestConstants.mMyInseadSuccessfulLoginTitle));
		Assert.assertEquals(driver.getTitle(), TestConstants.mMyInseadSuccessfulLoginTitle);	
	}

	// --------------------------------------------------------------------------------------------------------
	// Get the latest processed data sync ID for MyINSEAD
	public static int getLatestProcessedDataSyncID(WebDriver driver, TestParameters testParameters) 
	{
		// Assumption: We are logged in to MyINSEAD as a user that has Global Admin rights
		// Go to Sync monitor
		driver.get(testParameters.mURLMyInsead+testParameters.msyncMonitor);
		Assert.assertEquals(driver.getTitle(), "MyINSEAD - Sync Monitor", "Unable to find Sync monitor page title - you must be logged in as a Global Admin");

		// Get the latest processed data sync id
		WebElement dataSyncID = driver.findElement(By.xpath("(//tbody/tr/td)[1]"));
		
		return Integer.parseInt(dataSyncID.getText());
	}
	
	// --------------------------------------------------------------------------------------------------------
	// Make sure that the latest data sync for the given user says "Processed"
	public static void waitUntilDataSync(WebDriver driver, MyInseadUser userToBeSynced, int minimumDataID, TestParameters testParameters)
	{
		// Assumption: We are logged in to MyINSEAD as a user that has Global Admin rights
		// Go to Sync monitor
		driver.get(testParameters.mURLMyInsead+ testParameters.msyncMonitor);
		Assert.assertEquals(driver.getTitle(), "MyINSEAD - Sync Monitor", "Unable to find Sync monitor page title - you must be logged in as a Global Admin");

		// Get all cells that refer to the user being synced
		List<WebElement> syncObjectCells = driver.findElements(By.xpath("//td[contains(.,'" + userToBeSynced.mEMPLID + "')]"));
		for (int iCell = 0; iCell < syncObjectCells.size(); ++iCell)
		{
			WebElement syncObjectDataID = syncObjectCells.get(iCell).findElement(By.xpath("preceding-sibling::*[4]"));
			System.out.println(syncObjectDataID.getText());
			int dataIDThisRow = Integer.parseInt(syncObjectDataID.getText());
			if (dataIDThisRow < minimumDataID)
			{
				syncObjectCells.remove(iCell);
				--iCell;
			}
		}
		int maxWaitSeconds = 120, waitPerRefresh = 5;
		int currentWaitIter = 0, maxWaitIters = maxWaitSeconds / waitPerRefresh;
		// We must have at least one row with this name AND that has a dataID higher than the minimum. 
		while (syncObjectCells.isEmpty() && ++currentWaitIter <= maxWaitIters)
		{
			// Hit refresh
			driver.findElement(By.id("hrefRefresh")).click();
			// Wait for a few seconds
	    	driver.manage().timeouts().implicitlyWait(waitPerRefresh, TimeUnit.SECONDS);

	    	syncObjectCells = driver.findElements(By.xpath("//td[contains(.,'" + userToBeSynced.mPrimaryFirstName + "')]"));	
			for (int iCell = 0; iCell < syncObjectCells.size(); ++iCell)
			{
				WebElement syncObjectParent = syncObjectCells.get(iCell).findElement(By.xpath(".."));
				WebElement syncObjectDataID = syncObjectParent.findElement(By.xpath("(//td)[1]"));
				int dataIDThisRow = Integer.parseInt(syncObjectDataID.getText());
				if (dataIDThisRow < minimumDataID)
				{
					syncObjectCells.remove(iCell);
					--iCell;
				}
			}
		}
		
		// Go through all rows and make sure that they say "PROCESSED". If any row is not processed,
		// refresh the page until it does.
		boolean allRowsProcessed = false;
		currentWaitIter = 0;
		while (!allRowsProcessed && ++currentWaitIter <= maxWaitIters)
		{
			allRowsProcessed = true;
			for (WebElement syncObjectCell : syncObjectCells)
			{
				// Get the sibling of the cell and wait until it says "PROCESSED"
				WebElement syncObjectParent = syncObjectCell.findElement(By.xpath(".."));
				WebElement syncObjectIsProcessed = syncObjectParent.findElement(By.xpath("(//td)[2]"));
				boolean thisRowProcessed = syncObjectIsProcessed.getText().compareTo("PROCESSED") == 0;
				allRowsProcessed = allRowsProcessed & thisRowProcessed;
			}
			if (!allRowsProcessed)
			{
				// Hit refresh
				driver.findElement(By.id("hrefRefresh")).click();
				// Wait for a few seconds
		    	driver.manage().timeouts().implicitlyWait(waitPerRefresh, TimeUnit.SECONDS);
			}
		}
		// We may have run out of time, so make sure all rows say "PROCESSED"
		Assert.assertTrue(allRowsProcessed, "Unable to ensure MyINSEAD sync");
	}
}
