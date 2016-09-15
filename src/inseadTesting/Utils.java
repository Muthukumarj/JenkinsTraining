package inseadTesting;

import java.io.File;
import java.io.IOException;

import org.apache.commons.io.FileUtils;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;

public class Utils {

	// --------------------------------------------------------------------------------------------------------
	// Takes a screenshot and saves it in the screenshots folder with an updated name. For example,
	// if the test name is "TEST", will save it as "TEST_1" the first time it is called, "TEST_2" the
	// second time it is called, etc.
	public static void saveScreenshot(String testName, WebDriver driver) {
		
		// Find the next available filename for screenshots. O(n) but we do not 
		// expect to have a lot of screenshots per test.
		int maxScreenshotsPerTest = 100;
		int currentIndex = 1;
		String screenshotDirectory = "screenshots/";
		String fileName = screenshotDirectory + testName + "_" + currentIndex + ".png";
		
		while ((new File(fileName)).exists() && currentIndex < maxScreenshotsPerTest) {
			++currentIndex;
			fileName = screenshotDirectory + testName + "_" + currentIndex + ".png";
		}
		// If we have a maximum of 100 screenshots per test, then the file with _100 in its name will be overwritten.
		System.out.println("Saving to " + fileName);
		
		File scrFile = ((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
		// Now you can do whatever you need to do with it, for example copy somewhere
		try {
			scrFile.mkdirs();
			FileUtils.copyFile(scrFile, new File(fileName));
		} catch (IOException e) {
			// Eat exception, we do not want screenshots causing tests to fail.
			System.out.println("**** ERROR: Could not take screenshot");
			e.printStackTrace();
		}
	}
	
		 public static String toCamelCase(String inputString) {
	       String result = "";
	       if (inputString.length() == 0) {
	           return result;
	       }
	       char firstChar = inputString.charAt(0);
	       char firstCharToUpperCase = Character.toUpperCase(firstChar);
	       result = result + firstCharToUpperCase;
	       for (int i = 1; i < inputString.length(); i++) {
	           char currentChar = inputString.charAt(i);
	           char previousChar = inputString.charAt(i - 1);
	           if (previousChar == ' ') {
	               char currentCharToUpperCase = Character.toUpperCase(currentChar);
	               result = result + currentCharToUpperCase;
	           } else {
	               char currentCharToLowerCase = Character.toLowerCase(currentChar);
	               result = result + currentCharToLowerCase;
	           }
	       }
	       return result;
	   }
		 
		 public static void saveScreenshotRegression(String testName, WebDriver driver, String folder) {

				// Find the next available filename for screenshots. O(n) but we do not
				// expect to have a lot of screenshots per test.
				int maxScreenshotsPerTest = 100;
				int currentIndex = 1;
				String screenshotDirectory = "screenshots/CRUDRegression/" + folder;
				String fileName = screenshotDirectory + testName + "_" + currentIndex + ".png";

				while ((new File(fileName)).exists() && currentIndex < maxScreenshotsPerTest) {
					++currentIndex;
					fileName = screenshotDirectory + testName + "_" + currentIndex + ".png";
				}
				// If we have a maximum of 100 screenshots per test, then the file with
				// _100 in its name will be overwritten.
				System.out.println("Saving to " + fileName);

				File scrFile = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);
				// Now you can do whatever you need to do with it, for example copy
				// somewhere
				try {
					scrFile.mkdirs();
					FileUtils.copyFile(scrFile, new File(fileName));
				} catch (IOException e) {
					// Eat exception, we do not want screenshots causing tests to fail.
					System.out.println("**** ERROR: Could not take screenshot");
					e.printStackTrace();
				}
			}
}
