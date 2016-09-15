package inseadTesting;

import java.io.File;
import java.io.FileInputStream;
import java.lang.reflect.Constructor;
import java.lang.reflect.Method;
import java.nio.file.Paths;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.Assert;
import org.testng.IAnnotationTransformer;
import org.testng.annotations.ITestAnnotation;

// --------------------------------------------------------------------------------------------------------
// This class is responsible for turning tests on or off depending upon the contents of an excel file.
public class InseadAnnotationTransformer implements IAnnotationTransformer {

	// Hold a map of test class names to a flag determining if the test class should be enabled or not
	private Map mTestClassToEnabled = null;
	
	// --------------------------------------------------------------------------------------------------------
	@Override
	public void transform(ITestAnnotation annotation, Class testClass, Constructor testConstructor, Method testMethod) {
		
		// We have built the transformer to be called on a method only
		Assert.assertTrue(testMethod != null, "InseadAnnotationTransformer called, but not on a method");

		// If the map of test class names to enabled flags is not initialized, do so now.
		if (null == mTestClassToEnabled)
		{
			initializeTestClassToEnabledMap();
//			initializeTestClassToEnabledMap2();
		}
		
		String methodParentClass = testMethod.getDeclaringClass().getName();
		boolean isEnabled = false;
		if (mTestClassToEnabled.containsKey(methodParentClass)) {
			isEnabled = (boolean) mTestClassToEnabled.get(methodParentClass);
		}
		annotation.setEnabled(isEnabled);
	}
	
	// --------------------------------------------------------------------------------------------------------
	// Initializes the map of class names to enabled/disabled flags from an excel file
	private void initializeTestClassToEnabledMap()
	{
		mTestClassToEnabled = new HashMap();
		
		String executionPath = Paths.get("").toAbsolutePath().toString();
		String excelFullFilename = executionPath + "/src/inseadTesting/" + TestConstants.mTestDataExcelFile;
		try {
			FileInputStream fStream = new FileInputStream(new File(excelFullFilename));

			//Get the workbook instance for XLS file 
			Workbook workbook = new XSSFWorkbook(fStream);

			//Get first sheet from the workbook
			Sheet sheet = workbook.getSheetAt(0);

			// Iterate over all available rows
			DataFormatter formatter = new DataFormatter(); //creating formatter using the default locale
			int numRows = sheet.getLastRowNum();
			int iTestsToExecute = -1;
			for (int iRow = 0; iRow < numRows; ++iRow)
			{
				Row row = sheet.getRow(iRow);
				if (null == row) {
					continue;
				}
				Cell cell = row.getCell(0);
				if (null == cell) {
					continue;
				}
				String firstRowCellText = formatter.formatCellValue(cell);
				if (firstRowCellText != null && firstRowCellText.compareTo("#Test suites to execute") == 0)
				{
					iTestsToExecute = iRow;
				}
			}

			// We should have a list of test suites to execute
			Assert.assertTrue(iTestsToExecute >= 0, "Unable to find list of test suites to execute in Excel file");
			
			int iRow = iTestsToExecute;
			Row row = sheet.getRow(iRow);
			while (row != null)
			{
				String testName = formatter.formatCellValue(row.getCell(0));
				String testClassName = "";
				boolean isSuiteEnabled = (formatter.formatCellValue(row.getCell(1)).compareTo("YES") == 0);
				String sprintNo =  formatter.formatCellValue(row.getCell(2));
				
				if (sprintNo != null && sprintNo != "") {
					testClassName = "inseadTesting.Sprint_" + sprintNo + ".Test" + testName;

				} else {
					testClassName = "inseadTesting.Test" + testName;

				}
				// System.out.println("testClassName " + testClassName);

				mTestClassToEnabled.put(testClassName, isSuiteEnabled);

				++iRow;
				row = sheet.getRow(iRow);
			}
			
	        //------------------------------------------------------------------------------
            //Get value for the first sheet from the workbook - Regression
			if(workbook.getNumberOfSheets() > 1) { 
    			Sheet sheet2 = workbook.getSheetAt(1);
    
    			// Iterate over all available rows
    			int numRows2 = sheet2.getLastRowNum();
    			int iTestsToExecute2 = -1;
    			for (int iRow2 = 0; iRow2 < numRows2; ++iRow2)
    			{
    				Row row2 = sheet2.getRow(iRow2);
    				if (null == row2) {
    					continue;
    				}
    				Cell cell = row2.getCell(0);
    				if (null == cell) {
    					continue;
    				}
    				String firstRowCellText = formatter.formatCellValue(cell);
    				if (firstRowCellText != null && firstRowCellText.compareTo("#Test suites to execute") == 0)
    				{
    					iTestsToExecute2 = iRow2;
    				}
    			}
    			// List of test suites to execute in Regression tab
    			Assert.assertTrue(iTestsToExecute2 >= 0, "Unable to find list of test suites to execute in Excel file");
    			
    			int iRow2 = iTestsToExecute2;
    			Row row2 = sheet2.getRow(iRow2);
    			while (row2 != null)
    			{
    				String testName = formatter.formatCellValue(row2.getCell(0));
    				String testClassName = "";
    				boolean isSuiteEnabled = (formatter.formatCellValue(row2.getCell(1)).compareTo("YES") == 0);
    				String sprintNo =  formatter.formatCellValue(row2.getCell(2));
    				
    				if (sprintNo != null && sprintNo != "") {
    					testClassName = "inseadTesting." + sprintNo + ".Test" + testName;
    
    				} else {
    					testClassName = "inseadTesting.Test" + testName;
    				}
    				// System.out.println("testClassName " + testClassName);
    
    				mTestClassToEnabled.put(testClassName, isSuiteEnabled);
    
    				++iRow2;
    				row2 = sheet2.getRow(iRow2);
    			}
			}
			
			//------------------------------------------------------------------------------
			//Get value for the second sheet from the workbook - Regression
            if(workbook.getNumberOfSheets() > 2) { 
    			Sheet sheet3 = workbook.getSheetAt(2);
    
    			// Iterate over all available rows
    			int numRows3 = sheet3.getLastRowNum();
    			int iTestsToExecute3 = -1;
    			for (int iRow3 = 0; iRow3 < numRows3; ++iRow3)
    			{
    				Row row3 = sheet3.getRow(iRow3);
    				if (null == row3) {
    					continue;
    				}
    				Cell cell = row3.getCell(0);
    				if (null == cell) {
    					continue;
    				}
    				String firstRowCellText = formatter.formatCellValue(cell);
    				if (firstRowCellText != null && firstRowCellText.compareTo("#Test suites to execute") == 0)
    				{
    					iTestsToExecute3 = iRow3;
    				}
    			}
    			// List of test suites to execute in Regression tab
    			Assert.assertTrue(iTestsToExecute3 >= 0, "Unable to find list of test suites to execute in Excel file");
    			
    			int iRow3 = iTestsToExecute3;
    			Row row3 = sheet3.getRow(iRow3);
    			while (row3 != null)
    			{
    				String testName = formatter.formatCellValue(row3.getCell(0));
    				String testClassName = "";
    				boolean isSuiteEnabled = (formatter.formatCellValue(row3.getCell(1)).compareTo("YES") == 0);
    				String sprintNo =  formatter.formatCellValue(row3.getCell(2));
    				
    				if (sprintNo != null && sprintNo != "") {
    					testClassName = "inseadTesting." + sprintNo + ".Test" + testName;
    
    				} else {
    					testClassName = "inseadTesting.Test" + testName;
    				}
    				// System.out.println("testClassName " + testClassName);
    
    				mTestClassToEnabled.put(testClassName, isSuiteEnabled);
    
    				++iRow3;
    				row3 = sheet3.getRow(iRow3);
    			}
            }
			
			// Close out all the files
			workbook.close();
			fStream.close();
		} catch (Exception e) {
			System.out.println(e);
			Assert.assertTrue(false, "Unable to open excel file at " + excelFullFilename);
		}
	}
	
}
