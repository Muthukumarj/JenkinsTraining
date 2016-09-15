package inseadTesting;

import java.io.BufferedReader;
import java.io.FileReader;
import java.io.IOException;
import java.nio.file.Paths;

import org.testng.Assert;

// --------------------------------------------------------------------------------------------------------
// Class used to compare the contents of a BASELINE html file with the provided string
public class AssertHTMLContents {

	// --------------------------------------------------------------------------------------------------------
	// Compare the contents of a baseline html file with the provided string
	public static void compare(String baselineFilename, String actualHTML)
	{
		// Get the execution path for this project
		String executionPath = Paths.get("").toAbsolutePath().toString();
		String baselineFullFilename = executionPath + "/src/inseadTesting/baselines/" + baselineFilename;
		
		try {
			// Use the following when you want to save out the actual HTML as a baseline
//			PrintWriter writer = new PrintWriter(baselineFullFilename);
//			writer.write(actualHTML);
//			writer.close();

			// Assumption - The baseline HTML file is in a folder called "baselines" relative to the "inseadTesting" folder
			StringBuilder stringBuilder = new StringBuilder();
			try (BufferedReader br = new BufferedReader(new FileReader(baselineFullFilename))) {
			    for(String line; (line = br.readLine()) != null; ) {
					stringBuilder.append(line);
					stringBuilder.append("\n");
			    }
			}
			String expectedHTML = stringBuilder.toString();

			// Assert
			Assert.assertEquals(expectedHTML, actualHTML, "HTML contents do not match");
		} catch (IOException e) {
			// TODO Auto-generated catch block
			System.out.println(e);
			Assert.fail("Unable to open baseline file " + baselineFullFilename);
			e.printStackTrace();
		}
	}

}
