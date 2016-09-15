package inseadTesting;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.io.PrintWriter;
import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.util.Arrays;
import java.util.Collection;
import java.util.Comparator;
import java.util.List;
import java.util.Map;

import org.testng.IInvokedMethod;
import org.testng.IReporter;
import org.testng.IResultMap;
import org.testng.ISuite;
import org.testng.ISuiteResult;
import org.testng.ITestContext;
import org.testng.ITestNGMethod;
import org.testng.ITestResult;
import org.testng.collections.Lists;
import org.testng.internal.Utils;
import org.testng.log4testng.Logger;
import org.testng.xml.XmlSuite;

/**
 * Reported designed to render self-contained HTML top down view of a testing
 * suite. **** Modified the default TestNG reporter to have a data-inline="true"
 * attribute on all html elements. This will ensure that the CSS gets inlined by
 * the Jenkins Email-Ext plugin. Which in turn ensures that CSS styles are seen
 * correctly in web mail clients like GMail.
 */
public class InseadEmailableReporter implements IReporter {

	private enum CellType {
		Neutral, Pass, Fail
	}

	private static final Logger L = Logger.getLogger(InseadEmailableReporter.class);

	// ~ Instance fields ------------------------------------------------------

	private PrintWriter m_out;

	private Integer m_testIndex;

	private int m_methodIndex;

	// ~ Methods --------------------------------------------------------------

	/** Creates summary of the run */
	@Override
	public void generateReport(List<XmlSuite> xml, List<ISuite> suites, String outdir) {
		try {
			m_out = createWriter(outdir);
		} catch (IOException e) {
			L.error("output file", e);
			return;
		}
		startHtml(m_out);
		generatePassFailSummary(suites);
		generateSuiteSummaryReport(suites);
		generateFailedMethodDetailsReport(suites);
		endHtml(m_out);
		m_out.flush();
		m_out.close();
	}

	protected PrintWriter createWriter(String outdir) throws IOException {
		new File(outdir).mkdirs();
		return new PrintWriter(new BufferedWriter(new FileWriter(new File(outdir, "insead-emailable-report.html"))));
	}

	/** Generates a summary of the total passed/failed methods */
	protected void generatePassFailSummary(List<ISuite> suites) {

		// First, count the number of passed and failed tests
		int qty_pass_s = 0;
		int qty_fail = 0;
		for (ISuite suite : suites) {
			Map<String, ISuiteResult> tests = suite.getResults();
			for (ISuiteResult r : tests.values()) {
				ITestContext overview = r.getTestContext();
				int q = overview.getPassedTests().size();
				qty_pass_s += q;
				q = getMethodSet(overview.getFailedTests(), suite).size();
				qty_fail += q;
			}
		}

		// Display the passed tests as a SPAN element
		m_out.println("<h2>Summary</h2>");
		m_out.println("<p>");
		m_out.println(
				"<span style=\"font-size:24px;padding:10px;background-color:#00af00;color:#ffffff;border: 1px solid #505050\">Passed: "
						+ qty_pass_s + "</span>");

		// Spacer SPAN
		m_out.println("<span style=\"padding-left:20px\"></span>");

		// Display the failed tests as a SPAN element
		m_out.println(
				"<span style=\"font-size:24px;padding:10px;background-color:#af0000;color:#ffffff;border: 1px solid #505050\">Failed: "
						+ qty_fail + "</span>");
		m_out.println("</p>");
	}

	/** Generates a table with details of each failed method */
	protected void generateFailedMethodDetailsReport(List<ISuite> suites) {

		m_out.println("<br><h2>Failed test details</h2>");
		String titleCellStyle = "background-color: #cccccc; border-right: 1px solid #505050; padding-right:15px";
		String contentCellStyle = "font-family:monospace";

		for (ISuite suite : suites) {
			Map<String, ISuiteResult> tests = suite.getResults();
			for (ISuiteResult r : tests.values()) {
				ITestContext overview = r.getTestContext();
				IResultMap failedTests = overview.getFailedTests();
				for (ITestResult testResult : failedTests.getAllResults()) {

					m_out.println(
							"<table cellspacing=\"0\" and cellpadding=\"5\" style=\"border: 1px solid #505050;\">");

					// Print out the filename
					m_out.println("<tr>");
					m_out.println("<td style=\"" + titleCellStyle + "\">File</td>");
					m_out.print("<td style=\"" + contentCellStyle + "\">");
					ITestNGMethod testMethod = testResult.getMethod();
					String testClassName = testMethod.getTestClass().getName();
					// Strip out package name "inseadTesting" from class name
					String fileName = testClassName.substring("inseadTesting.".length()) + ".java";
					m_out.print(fileName);
					m_out.println("</td>");
					m_out.println("</tr>");

					// Print out the method name
					m_out.println("<tr>");
					m_out.println("<td style=\"" + titleCellStyle + "\">Method</td>");
					m_out.print("<td style=\"" + contentCellStyle + "\">");
					m_out.print(testMethod.getMethodName());
					m_out.println("</td>");
					m_out.println("</tr>");

					// Print out the line number in the file where the exception
					// occurred
					int iLineNumber = -1;
					for (StackTraceElement stackTraceElement : testResult.getThrowable().getStackTrace()) {
						if (0 == stackTraceElement.getMethodName().compareTo(testMethod.getMethodName())) {
							iLineNumber = stackTraceElement.getLineNumber();
							break;
						}
					}
					m_out.println("<tr>");
					m_out.println("<td style=\"" + titleCellStyle + "\">Line number</td>");
					m_out.print("<td style=\"" + contentCellStyle + "\">");
					m_out.print(iLineNumber);
					m_out.println("</td>");
					m_out.println("</tr>");

					if(testResult.getThrowable() != null && testResult.getThrowable().getMessage() != null) {
    					// Print out topmost stack trace element
    					String errorMessage = testResult.getThrowable().getMessage();
    					// Truncate the error message. For details the user can look
    					// at the build server
    					int maxChars = 200;
    					if (errorMessage.length() > maxChars) {
    						errorMessage = errorMessage.substring(0, maxChars) + " ...";
    					}
    					m_out.println("<tr>");
    					m_out.println("<td style=\"" + titleCellStyle + "\">Error details</td>");
    					m_out.print("<td style=\"" + contentCellStyle + "\">");
    					m_out.print(errorMessage);
    					m_out.println("</td>");
    					m_out.println("</tr>");
					}
					
					m_out.println("</table><br>");
				}
			}
		}
	}

	protected void generateExceptionReport(Throwable exception, ITestNGMethod method) {
		m_out.print("<div class=\"stacktrace\">");
		m_out.print(Utils.stackTrace(exception, true)[0]);
		m_out.println("</div>");
	}

	/**
	 * Since the methods will be sorted chronologically, we want to return the
	 * ITestNGMethod from the invoked methods.
	 */
	private Collection<ITestNGMethod> getMethodSet(IResultMap tests, ISuite suite) {
		List<IInvokedMethod> r = Lists.newArrayList();
		List<IInvokedMethod> invokedMethods = suite.getAllInvokedMethods();
		for (IInvokedMethod im : invokedMethods) {
			if (tests.getAllMethods().contains(im.getTestMethod())) {
				r.add(im);
			}
		}
		Arrays.sort(r.toArray(new IInvokedMethod[r.size()]), new TestSorter());
		List<ITestNGMethod> result = Lists.newArrayList();

		// Add all the invoked methods
		for (IInvokedMethod m : r) {
			result.add(m.getTestMethod());
		}

		// Add all the methods that weren't invoked (e.g. skipped) that we
		// haven't added yet
		for (ITestNGMethod m : tests.getAllMethods()) {
			if (!result.contains(m)) {
				result.add(m);
			}
		}
		return result;
	}

	public void generateSuiteSummaryReport(List<ISuite> suites) {
		tableStart("testOverview", null);
		m_out.print("<tr>");
		tableColumnStart("Suite");
		tableColumnStart("Passed");
		tableColumnStart("Failed");
		tableColumnStart("Execution rate %");
		tableColumnStart("Time (seconds)");
		m_out.println("</tr>");
		NumberFormat formatter = new DecimalFormat("#,##0.0");
		int qty_tests = 0;
		int qty_pass_s = 0;
		int qty_fail = 0;
		long time_start = Long.MAX_VALUE;
		long time_end = Long.MIN_VALUE;
		m_testIndex = 1;
		for (ISuite suite : suites) {
			if (suites.size() > 1) {
				titleRow(suite.getName(), 8);
			}
			Map<String, ISuiteResult> tests = suite.getResults();
			for (ISuiteResult r : tests.values()) {
				qty_tests += 1;
				ITestContext overview = r.getTestContext();

				int numPassedTests = overview.getPassedTests().size();
				int numFailedTests = getMethodSet(overview.getFailedTests(), suite).size();

				if (numPassedTests > 0 || numFailedTests > 0) {
					startSummaryRow(overview.getName());

					// Add passed tests
					qty_pass_s += numPassedTests;
					summaryCell(Integer.toString(numPassedTests),
							(numPassedTests > 0) ? CellType.Pass : CellType.Neutral);

					// Add failed tests
					qty_fail += numFailedTests;
					summaryCell(Integer.toString(numFailedTests),
							(numFailedTests > 0) ? CellType.Fail : CellType.Neutral);

					// Add percentage execution rate
					int numTotalMethods = overview.getAllTestMethods().length + overview.getExcludedMethods().size();
					double pctExecutionRate = (double) (numPassedTests + numFailedTests) / (numTotalMethods) * 100.0;
					summaryCell(Double.toString(pctExecutionRate), CellType.Neutral);

					// Add time taken to execute tests
					time_start = Math.min(overview.getStartDate().getTime(), time_start);
					time_end = Math.max(overview.getEndDate().getTime(), time_end);
					summaryCell(
							formatter.format(
									(overview.getEndDate().getTime() - overview.getStartDate().getTime()) / 1000.0),
							CellType.Neutral);
					m_out.println("</tr>");
				}

				m_testIndex++;
			}
		}
		if (qty_tests > 1) {
			m_out.println(
					"<tr style=\"background-color:#cccccc\"><td style=\"border: 1px solid #505050; padding: 5px 10px 5px 10px;\">Total</td>");
			summaryCell(Integer.toString(qty_pass_s), CellType.Neutral);
			summaryCell(Integer.toString(qty_fail), CellType.Neutral);
			summaryCell("", CellType.Neutral); // Dont add "total" for %
												// execution
			summaryCell(formatter.format((time_end - time_start) / 1000.), CellType.Neutral);
		}
		m_out.println("</table>");
	}

	private void summaryCell(String v, CellType cellType) {

		// The default colors below will be used for CellType.Neutral
		String backgroundColor = "#ffffff";
		String foregroundColor = "#000000";

		if (CellType.Pass == cellType) {
			backgroundColor = "#00af00";
			foregroundColor = "#ffffff";
		} else if (CellType.Fail == cellType) {
			backgroundColor = "#af0000";
			foregroundColor = "#ffffff";
		}

		m_out.print(
				"<td style=\"border: 1px solid #505050; padding: 5px 10px 5px 10px; text-align:center;background-color:"
						+ backgroundColor + ";color:" + foregroundColor + "\">" + v + "</td>");
	}

	private void startSummaryRow(String label) {
		m_out.print(
				"<tr>" + "<td  style=\"border: 1px solid #505050; padding: 5px 10px 5px 10px; text-align:left;padding-right:2em\">"
						+ label + "</td>");
	}

	private void tableStart(String cssclass, String id) {
		m_out.println("<br><table cellspacing=\"0\" cellpadding=\"0\""
				+ (cssclass != null ? " class=\"" + cssclass + "\"" : " style=\"padding-bottom:2em\"")
				+ (id != null ? " id=\"" + id + "\"" : "") + ">");
	}

	private void tableColumnStart(String label) {
		m_out.print(
				"<th style=\"border: 1px solid #505050; background-color: #cccccc; color:#505050;  padding: 5px 10px 5px 10px;\">"
						+ label + "</th>");
	}

	private void titleRow(String label, int cq) {
		titleRow(label, cq, null);
	}

	private void titleRow(String label, int cq, String id) {
		m_out.print("<tr");
		if (id != null) {
			m_out.print(" id=\"" + id + "\"");
		}
		m_out.println(
				"><th style=\"border: 1px solid #505050; background-color: #cccccc; color:#505050;  padding: 5px 10px 5px 10px;\" colspan=\""
						+ cq + "\">" + label + "</th></tr>");
	}

	/** Starts HTML stream */
	protected void startHtml(PrintWriter out) {
		out.println(
				"<!DOCTYPE html PUBLIC \"-//W3C//DTD XHTML 1.1//EN\" \"http://www.w3.org/TR/xhtml11/DTD/xhtml11.dtd\">");
		out.println("<html xmlns=\"http://www.w3.org/1999/xhtml\">");
		out.println("<head>");
		out.println("<title>TestNG Report</title>");
		out.println("</head>");
		out.println("<body style=\"font-family: Arial, Helvetica, sans-serif; font-size:14px\">");
		out.println("<h1>MyINSEAD SPRINT 53 test report</h1>");
	}

	/** Finishes HTML stream */
	protected void endHtml(PrintWriter out) {
		out.println("</body></html>");
	}

	// ~ Inner Classes --------------------------------------------------------
	/** Arranges methods by classname and method name */
	private static final class TestSorter implements Comparator<IInvokedMethod> {
		// ~ Methods
		// -------------------------------------------------------------

		/** Arranges methods by classname and method name */
		@Override
		public int compare(IInvokedMethod o1, IInvokedMethod o2) {
			return (int) (o1.getDate() - o2.getDate());
		}
	}
}
