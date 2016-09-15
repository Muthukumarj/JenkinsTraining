package inseadTesting;

import java.io.OutputStream;
import java.net.URL;
import javax.net.ssl.HttpsURLConnection;
import org.testng.ITestContext;
import org.testng.ITestResult;

public class FailureListener implements org.testng.ITestListener {

	@Override
	public void onTestStart(ITestResult result) {
		// Do nothing, we are only listening for failed tests
	}

	@Override
	public void onTestSuccess(ITestResult result) {
		// Do nothing, we are only listening for failed tests
	}

	@Override
	public void onTestFailure(ITestResult result) {

		// Log a defect in ServiceNow for the failed test
		String jobName = System.getenv("JOB_NAME");
		String buildNumber = System.getenv("BUILD_NUMBER");
		System.out.println("** Environment variables: " + jobName + ", " + buildNumber);
		if ((jobName == null)|| (buildNumber == null)) {
			System.out.println("**** Not logging a defect in ServiceNow as we are not running under Jenkins");
			return;
		}

		String className = result.getTestClass().getName();
		if (className.substring(0, 14).compareTo("inseadTesting.") == 0) {
			// Remove the "inseadTesting." from the start of the string
			className = className.substring(14);
		}
		String defectDescription = "Test Failure: " + jobName + ", Build# " + buildNumber + " " + className + "." + result.getMethod().getMethodName();
		
		try {
			String url = "https://inseadtest.service-now.com/api/now/v1/table/rm_defect";
			URL obj = new URL(url);
			HttpsURLConnection con = (HttpsURLConnection) obj.openConnection();
	
			//add request header
			con.setRequestMethod("POST");
			con.setRequestProperty("Accept", "application/json");
			con.setRequestProperty("Authorization", "Basic amVua2luc19hcGlfaW50ZWdyYXRpb246ZGVmZWN0NzAwMg==");
			con.setRequestProperty("Content-type", "application/json");
			con.setRequestProperty("x-usertoken", "471071960f335a00f43283fc22050e9bb4b55f4059c51544b9a646045cf34873a56770b5");
			con.setDoOutput(true);
			
			// set the short description as a JSON string like "{'short_description': 'Defect Description'}"
			StringBuilder paramsBuilder = new StringBuilder();
			paramsBuilder.append("{");
			paramsBuilder.append("'short_description': '" + defectDescription + "'");
			paramsBuilder.append(",");
			paramsBuilder.append("'cmdb_ci': '28f6e3692bf96500f4358c1fe8da155d'");
			paramsBuilder.append("}");
			byte[] shortDescriptionBytes = paramsBuilder.toString().getBytes();
			OutputStream os = con.getOutputStream();
			os.write(shortDescriptionBytes);
			os.close();
			
	
			// Debugging code to get the response from the server. Uncomment for debugging.
//			int responseCode = con.getResponseCode();
//			System.out.println("Raising a ServiceNow defect. Response Code : " + responseCode);
//			System.out.println("Raising a ServiceNow defect. Response Message : " + con.getResponseMessage());
//			BufferedReader in = new BufferedReader(new InputStreamReader(con.getInputStream()));
//			String inputLine;
//			StringBuffer response = new StringBuffer();
//			while ((inputLine = in.readLine()) != null) {
//				response.append(inputLine);
//			}
//			in.close();
////			Print result
//			System.out.println(response.toString());
			
		} catch (Exception e) {
			// Eating exception here. We don't want a failure to create defects to cause the test reporting to fail.
			System.out.println("**** ERROR while trying to raise a defect in ServiceNow");
			e.printStackTrace();
		}
		
	}

	@Override
	public void onTestSkipped(ITestResult result) {
		// Do nothing, we are only listening for failed tests
	}

	@Override
	public void onTestFailedButWithinSuccessPercentage(ITestResult result) {
		// Do nothing, we are only listening for failed tests
	}

	@Override
	public void onStart(ITestContext context) {
		// Do nothing, we are only listening for failed tests
	}

	@Override
	public void onFinish(ITestContext context) {
		// Do nothing, we are only listening for failed tests
	}

}
