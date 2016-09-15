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
	  driver.get("https://iconnect.insead.edu/Pages/homepage.aspx");		  
		
	  WebElement myProfileLink = driver.findElement(By.linkText("Events Calendar"));		  
	  
	  Assert.assertNotNull(myProfileLink);

	  myProfileLink.click();		  		  
  
	  // Wait until the tab has loaded
	  WebDriverWait wait = new WebDriverWait(driver, 15);
	  wait.until(ExpectedConditions.titleContains("iConnect Global Events"));

	  WebElement result = driver.findElement(By.id("ctl00_SPWebPartManager1_g_1cb1fb74_d22d_4a35_9c14_8d0ad940656b_ctl00_lblCalLabel"));
  
	  Assert.assertEquals(result.getText().toLowerCase(), "september,2016"); 

  }
}
