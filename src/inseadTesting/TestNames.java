package inseadTesting;

import java.net.URL;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.testng.annotations.Test;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.Parameters;
import org.testng.Assert;
import org.openqa.selenium.By;
import org.openqa.selenium.Platform;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.remote.RemoteWebDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

public class TestNames {

  private WebDriver driver = null;
  
  // --------------------------------------------------------------------------------------------------------
  // Create a browser before every test method
  @Parameters({ "server-host", "server-port", "browser-name", "browser-version", "operating-system" })
  @BeforeMethod(alwaysRun = true)
  public void beforeMethod(String serverHost, String serverPort, String browserName, String browserVersion, String operatingSystem) {

	  try {
		  // Construct the string used to connect to the remote web driver (servername+port)
		  String connectionString = "http://" + serverHost + ":" + serverPort + "/wd/hub";
		  DesiredCapabilities capability = new DesiredCapabilities(browserName, browserVersion, Platform.extractFromSysProperty(operatingSystem));
		  driver = new RemoteWebDriver(new URL(connectionString), capability);
	  } catch (Exception e) {
		  System.out.println(e);
		  throw new IllegalStateException("Can't start web driver", e);
	  }
  }

  // --------------------------------------------------------------------------------------------------------
  // Destroy the browser after every test method
  @AfterMethod(alwaysRun = true)
  public void afterMethod() {
	  driver.close();
	  driver.quit();
	  driver = null;
  }
  
  // --------------------------------------------------------------------------------------------------------
  // Update an active Primary Name in Peoplesoft
  @Test(dataProvider = "excelData", dataProviderClass = ExcelDataProvider.class)
  public void FUNC321002_1_Update_an_active_Primary_Name_in_Peoplesoft(TestParameters testParameters, MyInseadUser user) throws InterruptedException
  {
	  // First, log in to MyINSEAD and get the latest processed data sync ID
	  LoginHelperMyInsead.loginViaLoginPage(driver, testParameters);
	  int latestProcessedDataSyncID = LoginHelperMyInsead.getLatestProcessedDataSyncID(driver, testParameters);

	  // Log in to account
	  PeoplesoftHelper.doLogin(driver, testParameters);
	  
	  // Make the primary name ready to be edited
	  int primaryRowIndex = makePrimaryNameEditable(driver, user);
	  
	  // Click on the primary name
	  driver.findElement(By.id("SCC_NAME_TYPE_LNK$" + primaryRowIndex)).click();
	  WebDriverWait wait = new WebDriverWait(driver, 15);
	  wait.until(ExpectedConditions.textToBePresentInElementLocated(By.id("DERIVED_SCC_NM_NAME_TYPE"), "Primary"));
	  wait.until(ExpectedConditions.textToBePresentInElementValue(By.id("DERIVED_SCC_NM_FIRST_NAME"), user.mPrimaryFirstName));
	  wait.until(ExpectedConditions.textToBePresentInElementValue(By.id("DERIVED_SCC_NM_LAST_NAME"), user.mPrimaryLastName));

	  // Change the primary name
	  user.mPrimaryFirstName += Math.round(1000 * Math.random());
	  user.mPrimaryLastName += Math.round(1000 * Math.random());
	  driver.findElement(By.id("DERIVED_SCC_NM_FIRST_NAME")).clear();
	  driver.findElement(By.id("DERIVED_SCC_NM_FIRST_NAME")).sendKeys(user.mPrimaryFirstName);
	  driver.findElement(By.id("DERIVED_SCC_NM_LAST_NAME")).clear();
	  driver.findElement(By.id("DERIVED_SCC_NM_LAST_NAME")).sendKeys(user.mPrimaryLastName);
	  driver.findElement(By.id("DERIVED_SCC_NM_SCC_ADDR_SUBMT_BTN")).click();
	  driver.findElement(By.id("#ICSave")).click();
	  driver.switchTo().defaultContent();
	  
	  // Execute the IN/OUT processes in PeopleSoft so that the data will be updated.
	  PeoplesoftHelper.syncDataPeoplesoftToMyINSEAD(driver, user);
	  
	  // Sign out of Peoplesoft
//	  driver.switchTo().defaultContent();
//	  driver.findElement(By.id("pthdr2logout")).click();
//	  wait.until(ExpectedConditions.titleIs(TestConstants.mPeoplesoftLoginPageTitle));
//	  
	  // Go to MyInsead and make sure that the data synchronization is complete
	  // Note that we are logging out and back in. Sometimes the "Login As User" button was not showing up if we used the same session.
	  driver.get(testParameters.mURLMyInsead);
	  wait.until(ExpectedConditions.titleIs(TestConstants.mMyInseadSuccessfulLoginTitle));
	  driver.findElement(By.linkText("Logout")).click();
	  wait.until(ExpectedConditions.titleIs(TestConstants.mMyInseadSuccessfulLogoutTitle));
	  LoginHelperMyInsead.loginViaLoginPage(driver, testParameters);
	  LoginHelperMyInsead.waitUntilDataSync(driver, user, latestProcessedDataSyncID + 1, testParameters);
	  latestProcessedDataSyncID++;
	  driver.get(testParameters.mURLMyInsead);
	  
	  // Go to MyInsead and make sure that the name is displayed correctly in the "My Profile" tab
	  LoginHelperMyInsead.loginViaGlobalAdmin(driver, user);

	  // Go to the "My Profile" tab
	  WebElement myProfileLink = driver.findElement(By.linkText("My Profile"));
	  myProfileLink.click();
	  
	  // Wait until the tab has loaded
	  wait.until(ExpectedConditions.visibilityOfElementLocated(By.className("profile_basic_details")));
	  
	  // Check to make sure the changed name is reflected
	  WebElement profileDetails = driver.findElement(By.className("profile_basic_details"));
	  WebElement nameField = profileDetails.findElement(By.xpath("(.//div)[1]"));
	  Assert.assertEquals(nameField.getText(), user.mPrimaryFirstName + " " + user.mPrimaryLastName);

	  // Revert the changes we have made for this name.
	  PeoplesoftHelper.doLogin(driver, testParameters);
	  makePrimaryNameEditable(driver, user);
	  driver.findElement(By.id("#ICSave")).click();
	  // Execute the IN/OUT processes in PeopleSoft so that the data will be updated.
	  driver.switchTo().defaultContent();
	  PeoplesoftHelper.syncDataPeoplesoftToMyINSEAD(driver, user);

	  // Go to MyInsead and make sure that the data synchronization is complete
	  // Note that we are logging out and back in. Sometimes the "Login As User" button was not showing up if we used the same session.
	  driver.get(testParameters.mURLMyInsead);
	  wait.until(ExpectedConditions.titleIs(TestConstants.mMyInseadSuccessfulLoginTitle));
	  driver.findElement(By.linkText("Logout")).click();
	  wait.until(ExpectedConditions.titleIs(TestConstants.mMyInseadSuccessfulLogoutTitle));
	  LoginHelperMyInsead.loginViaLoginPage(driver, testParameters);
	  LoginHelperMyInsead.waitUntilDataSync(driver, user, latestProcessedDataSyncID + 1, testParameters);
	  latestProcessedDataSyncID++;
  }
  
  // --------------------------------------------------------------------------------------------------------
  // Makes a primary name editable by removing all history entries for today
  private int makePrimaryNameEditable(WebDriver driver, MyInseadUser user)
  {
	  // Go to the screen where we can change names
	  WebDriverWait wait = new WebDriverWait(driver, 15);
	  wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("pthnavbca_PORTAL_ROOT_OBJECT")));
	  driver.findElement(By.id("pthnavbca_PORTAL_ROOT_OBJECT")).click();
	  wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("fldra_HCCC_BUILD_COMMUNITY")));
	  driver.findElement(By.id("fldra_HCCC_BUILD_COMMUNITY")).click();
	  wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("fldra_HCCC_PERSONAL_INFORMATION")));
	  driver.findElement(By.id("fldra_HCCC_PERSONAL_INFORMATION")).click();
	  wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("fldra_HCCC_BIOGRAPHICAL")));
	  driver.findElement(By.id("fldra_HCCC_BIOGRAPHICAL")).click();
	  wait.until(ExpectedConditions.visibilityOfElementLocated(By.linkText("Names")));
	  driver.findElement(By.linkText("Names")).click();
	  
	  // Search for the employee to change the name of
	  driver.switchTo().frame("TargetContent");
	  wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("PEOPLE_SRCH_EMPLID")));
	  driver.findElement(By.id("PEOPLE_SRCH_EMPLID")).clear();
	  driver.findElement(By.id("PEOPLE_SRCH_EMPLID")).sendKeys(user.mEMPLID);
	  driver.findElement(By.id("#ICSearch")).click();
	  wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("win0div$ICField1")));
	  Assert.assertEquals(driver.findElement(By.id("win0div$ICField1")).getText(), "Names");
	  
	  // Click on the "Correct history" field
	  driver.findElement(By.id("#ICCorrection")).click();
	  wait.until(ExpectedConditions.invisibilityOfElementLocated(By.id("WAIT_win0")));	// Wait till the "processing/wait" spinner dissapears from the top right
	  
	  // Go to the Name history and make sure that no changes have been made today. If changes have been made, delete them.
	  // This is because a primary name can be changed only once per day.
	  int maxRows = 100;	// Unlikely we will have more rows than this
	  int currentRow = 0;
	  WebElement nameTypeElement = driver.findElement(By.id("SCC_NAME_TYPE_LNK$" + currentRow));
	  while (nameTypeElement != null && nameTypeElement.getText().compareTo("Primary") != 0 && currentRow <= maxRows) 
	  {
		  ++currentRow;
		  nameTypeElement = driver.findElement(By.id("SCC_NAME_TYPE_LNK$" + currentRow));
	  }
	  Assert.assertTrue(nameTypeElement.getText().compareTo("Primary") == 0);
	  int primaryRowIndex = currentRow;
	  
	  // Click on "Name History"
	  driver.findElement(By.id("DERIVED_SCC_NM_NAME_HISTORY_BTN$" + currentRow)).click();
	  wait.until(ExpectedConditions.textToBePresentInElementLocated(By.id("app_label"), "Name Type History"));
	  
	  currentRow = 0;
	  String todaysDate = new SimpleDateFormat("dd/MM/yyyy").format(new Date());
	  boolean rowExists = driver.findElements(By.id("SCC_NAMES_L1_H_EFFDT$" + currentRow)).size() > 0;
	  while (rowExists)
	  {
		  // This row exists. Get date for this row
		  WebElement effectiveDateElement = driver.findElement(By.id("SCC_NAMES_L1_H_EFFDT$" + currentRow));
		  String dateInRow = effectiveDateElement.getAttribute("value");
		  if (dateInRow.compareTo(todaysDate) == 0)
		  {
			  // This row refers to todays date. Delete the row.
			  String originalRowCountText = driver.findElement(By.className("PSGRIDCOUNTER")).getText();
			  int numRecords = Integer.parseInt(originalRowCountText.substring(originalRowCountText.indexOf('-') + 1, originalRowCountText.indexOf(' ')));
			  String newRowCountText = "1-" + (numRecords - 1) + " of " + (numRecords - 1);
			  driver.findElement(By.id("$ICField36$delete$0$$" + currentRow)).click();
			  driver.switchTo().defaultContent();
			  wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("#ALERTOK")));
			  driver.findElement(By.id("#ALERTOK")).click();
			  wait.until(ExpectedConditions.invisibilityOfElementLocated(By.id("#ALERTOK")));
			  driver.switchTo().frame("TargetContent");
			  wait.until(ExpectedConditions.textToBePresentInElementLocated(By.className("PSGRIDCOUNTER"), newRowCountText));
			  --currentRow;
		  }
		  ++currentRow;
		  rowExists = driver.findElements(By.id("SCC_NAMES_L1_H_EFFDT$" + currentRow)).size() > 0;
	  }
	  driver.findElement(By.id("#ICSave")).click();
	  wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("win0div$ICField1")));
	  Assert.assertEquals(driver.findElement(By.id("win0div$ICField1")).getText(), "Names");
	  
	  return primaryRowIndex;
  }
}
