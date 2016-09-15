package inseadTesting;

import java.io.File;

import org.apache.commons.io.FileUtils;
import org.openqa.selenium.By;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Assert;

//--------------------------------------------------------------------------------------------------------
// Class to help with the "My Profile" tab in MyInsead
public class MyProfileHelper {

	//--------------------------------------------------------------------------------------------------------
	// Set the given users location in MyInsead. Assumes we are logged in.
	public static void setLocation(WebDriver driver, MyInseadUser user) 
	{
		// Go to the "My Profile" tab
		WebElement myProfileLink = driver.findElement(By.linkText("My Profile"));
		myProfileLink.click();
		  
		// Wait until the tab has loaded
		WebDriverWait wait = new WebDriverWait(driver, 15);
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("myProfile")));

		// Click on the button to edit the location
		WebElement editLocationButton = driver.findElement(By.id("locationFormLink"));
		editLocationButton.click();

		// Wait till the location form is visible
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("form_country")));

		// Select the given location as the users location
		Select countryDropdown = new Select(driver.findElement(By.id("form_country")));
		countryDropdown.selectByVisibleText(user.mLocation);
		
		// Click Save
		WebElement saveLocationButton = driver.findElement(By.xpath("//input[@value='Save']"));
		saveLocationButton.click();

		// Wait until the form is hidden
		wait.until(ExpectedConditions.invisibilityOfElementLocated(By.id("form_country")));		
		// Wait until the location element is updated
		wait.until(ExpectedConditions.textToBePresentInElementLocated(By.xpath("//*[@id='locationFormLink']/.."), user.mLocation));
		
		// Verify that the location has been updated
		editLocationButton = driver.findElement(By.id("locationFormLink"));
		WebElement locationDiv = editLocationButton.findElement(By.xpath(".."));	// This is the parent div of "editLocationButton"
		Assert.assertEquals(locationDiv.getText(), user.mLocation, "User location was not set correctly");
	}
		
		public static void takescreenshotGeo(WebDriver driver, int row, String emplid, Boolean result) throws Exception
	       {
	              if(result==true)
	                {
	                     TakesScreenshot ts = ((TakesScreenshot)driver);
	                     File ss = ts.getScreenshotAs(OutputType.FILE);
	              FileUtils.copyFile(ss, new File("C:/Users/J.GARCIA/github/myINSEAD/src/Geoloc/Screenshot/"+"Emplid_"+row+"-"+emplid+"_Global Admin-Passed.png"));
	                }
	                else{ 
	                       TakesScreenshot ts = ((TakesScreenshot)driver);
	                           File ss = ts.getScreenshotAs(OutputType.FILE);
	                           FileUtils.copyFile(ss, new File("C:/Users/J.GARCIA/github/myINSEAD/src/Geoloc/Screenshot/"+"Emplid_"+row+"-"+emplid+"_Global Admin-Failed.png"));
	                }    
	       }
	       
	       public static void takescreenshotGeoMyProfile(WebDriver driver, int row, String emplid) throws Exception
	       {
	              
	                       TakesScreenshot ts = ((TakesScreenshot)driver);
	                           File ss = ts.getScreenshotAs(OutputType.FILE);
	                           FileUtils.copyFile(ss, new File("C:/Users/J.GARCIA/github/myINSEAD/src/Geoloc/Screenshot/"+"Emplid_"+row+"-"+emplid+"-MyProfile.png"));
	                
	       }
	       
	       public static void takescreenshotnoGeo(WebDriver driver, int row, String emplid, String result) throws Exception{
	              TakesScreenshot ts = ((TakesScreenshot)driver);
	              File ss = ts.getScreenshotAs(OutputType.FILE);
	       FileUtils.copyFile(ss, new File("C:/Users/J.GARCIA/github/myINSEAD/src/Geoloc/Screenshot/"+"Emplid_"+row+"-"+emplid+"-"+ result +".png"));
	}
	       
	       
	       public static void takescreenshotrememberUsername(WebDriver driver, int tc) throws Exception{
	              TakesScreenshot ts = ((TakesScreenshot)driver);
	              File ss = ts.getScreenshotAs(OutputType.FILE);
	       FileUtils.copyFile(ss, new File("C:/Users/J.GARCIA/github/myINSEAD/src/RememberUsername/Screenshot/"+"TC "+tc +"-Passed.png"));
	}
	       
	       public static void takescreenshotWETC1(WebDriver driver) throws Exception
	       {
	              
	                       TakesScreenshot ts = ((TakesScreenshot)driver);
	                           File ss = ts.getScreenshotAs(OutputType.FILE);
	                           FileUtils.copyFile(ss, new File("C:/Users/J.GARCIA/github/myINSEAD/src/WorkExperience/Screenshot/"+"WETC1_Note_in_Business_Address.png"));
	                
	}
			
	       public static void takescreenshotWETC2(WebDriver driver) throws Exception
	       {
	              
	                       TakesScreenshot ts = ((TakesScreenshot)driver);
	                           File ss = ts.getScreenshotAs(OutputType.FILE);
	                           FileUtils.copyFile(ss, new File("C:/Users/J.GARCIA/github/myINSEAD/src/WorkExperience/Screenshot/"+"WETC2_Note_in_Business_Address.png"));
	                
	       }
	       

	       public static void takescreenshotWETC3(WebDriver driver) throws Exception
	       {
	              
	                       TakesScreenshot ts = ((TakesScreenshot)driver);
	                           File ss = ts.getScreenshotAs(OutputType.FILE);
	                           FileUtils.copyFile(ss, new File("C:/Users/J.GARCIA/github/myINSEAD/src/WorkExperience/Screenshot/"+"WETC3_BusinessAddress_fields_readonly.png"));
	                
	       }
}
