package inseadTesting;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Assert;

//--------------------------------------------------------------------------------------------------------
// Class to help with the "My Association" tab in MyINSEAD
public class MyAssociationHelper {

	//--------------------------------------------------------------------------------------------------------
	// Sets the country in MyAssociation to Finland
	public static void setAssociationCountryFinland(WebDriver driver, MyInseadUser user)
	{
		  // Go to the My Associations link
		  WebElement myAssociationsLink = driver.findElement(By.linkText("My Association"));
		  myAssociationsLink.click();
		  
		  // Wait until the tab has loaded
		  WebDriverWait wait = new WebDriverWait(driver, 15);
		  wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("divMembership")));

		  // Check to see that header has the right text
		  WebElement divMembership = driver.findElement(By.id("divMembership"));
		  WebElement membershipTitle = divMembership.findElement(By.xpath(".//h2"));
		  Assert.assertEquals(membershipTitle.getText(), "Membership", "Missing text 'Membership'");
		  
		  // Check to see that the country <span> has the correct text
		  WebElement membershipCountry = divMembership.findElement(By.xpath(".//span"));
		  Assert.assertEquals(membershipCountry.getText(), user.mLocation, "The users location on the Membership page was not populated correctly");
		  
		  // Check to see if the links are correct
		  WebElement joinLink = driver.findElement(By.className("hrefMembershipPopup"));
		  Assert.assertTrue(joinLink != null, "Unable to find the link for joining a group");
		  Assert.assertEquals(joinLink.getText(), "here", "Link to join membership group has incorrect text");
		  WebElement changeLocationLink = driver.findElement(By.id("hrefChooseLocation"));
		  Assert.assertTrue(changeLocationLink != null, "Unable to find the link for changing location");
		  Assert.assertEquals(changeLocationLink.getText(), "here", "Link to change location has incorrect text");
		  
		  // Also check to make sure that the entire text in the div is correct
		  String expectedMembershipString = "Membership\n";
		  expectedMembershipString += "Join your local National Alumni Association to receive exclusive content and privileges.\n\n";
		  expectedMembershipString += "Our records say that you are in " + user.mLocation + ". Click here to join the INSEAD Alumni Association " + user.mLocation + ", or here to choose another location.";
		  String actualMembershipString = divMembership.getText();
		  Assert.assertEquals(actualMembershipString, expectedMembershipString, "The Membership page contents were not as expected");
		  
		  // Click to change your location to Finland
		  changeLocationLink.click();
		  wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("divAssociationFieldset")));
		  wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//a[text()='Finland']")));
		  WebElement finlandLink = driver.findElement(By.xpath("//a[text()='Finland']"));
		  finlandLink.click();
		  
		  // Wait until the product list is shown
		  wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("people_list")));

		  // Now check to make sure products are seen as expected. Get HTML contents of the "divProductList"
		  WebElement divProductList = driver.findElement(By.id("divProductList"));
		  String divProductHTML = divProductList.getAttribute("innerHTML");
		  // Compare the HTML contents with a saved baseline
//		  AssertHTMLContents.compare("finland.html", divProductHTML);
	}

	//--------------------------------------------------------------------------------------------------------
	// Populates the billing page with the given data
	public static void populateBillingPage(WebDriver driver, MyInseadUser user)
	{
		  driver.findElement(By.id("billing:firstname")).clear();
		  driver.findElement(By.id("billing:firstname")).sendKeys(user.mPrimaryFirstName);

		  driver.findElement(By.id("billing:lastname")).clear();
		  driver.findElement(By.id("billing:lastname")).sendKeys(user.mPrimaryLastName);

		  driver.findElement(By.id("billing:company")).clear();
		  driver.findElement(By.id("billing:company")).sendKeys(user.mCompany);

		  driver.findElement(By.id("billing:email")).clear();
		  driver.findElement(By.id("billing:email")).sendKeys(user.mBillingEmailAddress);

		  driver.findElement(By.id("billing:street1")).clear();
		  driver.findElement(By.id("billing:street1")).sendKeys(user.mStreetAddress1);
		  
		  driver.findElement(By.id("billing:street2")).clear();
		  driver.findElement(By.id("billing:street2")).sendKeys(user.mStreetAddress2);
		  
		  driver.findElement(By.id("billing:street3")).clear();
		  driver.findElement(By.id("billing:street3")).sendKeys(user.mStreetAddress3);
		  
		  driver.findElement(By.id("billing:city")).clear();
		  driver.findElement(By.id("billing:city")).sendKeys(user.mCity);
		  
		  driver.findElement(By.id("billing:region")).clear();
		  driver.findElement(By.id("billing:region")).sendKeys(user.mStateProvince);
		  
		  driver.findElement(By.id("billing:postcode")).clear();
		  driver.findElement(By.id("billing:postcode")).sendKeys(user.mZipPostalCode);

		  new Select(driver.findElement(By.id("billing:country_id"))).selectByVisibleText(user.mBillingCountry);

		  driver.findElement(By.id("billing:telephone")).clear();
		  driver.findElement(By.id("billing:telephone")).sendKeys(user.mTelephone);

		  driver.findElement(By.id("billing:fax")).clear();
		  driver.findElement(By.id("billing:fax")).sendKeys(user.mFax);
	}
}
