package utils;

import java.util.Date;

import org.apache.log4j.LogManager;
import org.apache.log4j.Logger;

import com.aventstack.extentreports.ExtentReports;
import com.aventstack.extentreports.ExtentTest;
import com.aventstack.extentreports.reporter.ExtentHtmlReporter;
import com.aventstack.extentreports.reporter.configuration.Theme;

public class ExtentReportClass {
	
	public static ExtentHtmlReporter htmlReporter;
	public static ExtentReports extentReportObject;
	public static ExtentTest extentTest;
	
	static String reportLocation = "/SeleniumCucumberBDDFramework/reports";
	static String reportDocumentTitle = "Automation Execution Summary";
	static String reportName = "Automation Report";
	
	public static Logger logger = LogManager.getLogger(ExtentReportClass.class);
	
	public static void initializeExtentReport() {
		
		String path = reportLocation + "ExecutionReportExtent.html";
		
		logger.info("INFO : Extent report Path = " + path);
		
		htmlReporter = new ExtentHtmlReporter(reportLocation);
		htmlReporter.config().setEncoding("utf-8");
		htmlReporter.config().setDocumentTitle(reportDocumentTitle);
		htmlReporter.config().setReportName(reportName);
		htmlReporter.config().setTheme(Theme.DARK);
		htmlReporter.setAppendExisting(true);
		
		extentReportObject = new ExtentReports();
		
		extentReportObject.attachReporter(htmlReporter);
		
		logger.info("PASS : Extent report object created successfully = " + extentReportObject);
		
	}
	
	public static ExtentTest InitializeExecutionReportUsingExtent(String scenarioOutline) {
		
		try {

			extentTest = extentReportObject.createTest(scenarioOutline);
			
			Date startTime = new Date();
			
			extentTest.getModel().setStartTime(startTime);
			
			logger.info("PASS : extentTest object created successfully = " + extentTest);
			
		}
		catch (Exception e) {
			logger.info("FAIL : Error in InitializeExecutionReportUsingExtent");
		}

		return extentTest;
		
	}
	
	public static void CleanUpExecutionReportUsingExtent(ExtentTest extentTest) {
		
		try {
			
			Thread.sleep(5000);
			
			Date endTime = new Date();
			
			extentTest.getModel().setEndTime(endTime);
			
			extentReportObject.flush();
			
		}
		catch (Exception e) {
			logger.info("FAIL : Error in CleanUpExecutionReportUsingExtent");
		}
		
	}

}
