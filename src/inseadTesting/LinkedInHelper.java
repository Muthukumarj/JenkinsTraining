package inseadTesting;

import java.util.List;

import org.openqa.selenium.By;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Assert;

// class to help LinkedIn Authentication 
public class LinkedInHelper {

	public static void doLogin(WebDriver driver, TestParameters testParameters, String screenshotName) 
	{

		WebDriverWait wait = new WebDriverWait(driver, 15);
		wait.until(ExpectedConditions.titleIs(TestConstants.linkedInLoginPageTitle));
		Assert.assertEquals(driver.getTitle(), TestConstants.linkedInLoginPageTitle);

		// Populate the username field
		WebElement usernameTextbox = driver.findElement(By.id("session_key-oauthAuthorizeForm"));
		usernameTextbox.clear();
		usernameTextbox.sendKeys(testParameters.linkedInUserName);

		// Populate the password field
		WebElement passwordTextbox = driver.findElement(By.id("session_password-oauthAuthorizeForm"));
		passwordTextbox.clear();
		passwordTextbox.sendKeys(testParameters.linkedInPassword);
		
		// take screen shot 
		
		 Utils.saveScreenshot(screenshotName, driver);

		// Click the sign in button
		WebElement signinButton = driver.findElement(By.name("authorize"));
		signinButton.click();

		
	}
	
	public static void completeForm(WebDriver driver, TestParameters testParameters) throws InterruptedException{
		
		WebDriverWait wait = new WebDriverWait(driver, 15);
		
		
		// Wait till the Industry is visible
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("form_industryId")));

		// fill Industry
		Select industryDropdown = new Select(driver.findElement(By.id("form_industryId")));
		industryDropdown.selectByVisibleText(testParameters.industryForJobForm);
		
		
		// Wait till the country is visible
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("form_country")));

		// fill country
		Select countryDropdown = new Select(driver.findElement(By.id("form_country")));
		countryDropdown.selectByVisibleText(testParameters.countryForJobForm);
		
		
		// fill company 
		driver.findElement(By.id("form_organisation")).clear();
	    driver.findElement(By.id("form_organisation")).sendKeys(testParameters.companyForJobForm);
	    Thread.sleep(3000);
	    List <WebElement> listItems = driver.findElements(By.xpath("//li[@class='ui-menu-item']/a"));
	    listItems.get(0).click();
	   
	}
	
	
	
}
