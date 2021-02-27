package utils;

import org.openqa.selenium.WebDriver;

import org.testng.Assert;

import com.aventstack.extentreports.ExtentTest;

import java.io.File;
import java.util.LinkedHashMap;

import org.apache.log4j.LogManager;
import org.apache.log4j.Logger;

public class InitScriptClass {
	
	ReadExcel excel = new ReadExcel();
	
	public static Logger logger = LogManager.getLogger(InitScriptClass.class);
	
	public long threadId = Thread.currentThread().getId();
	
	ApplicationPropertiesReader appReader = new ApplicationPropertiesReader();
	
	ConfigReader configReader = new ConfigReader();
	
	public String tagName = "";
	public String sheetName = "";
	
	public static ExtentTest currentlyExecutedTest;
	
	public static WebDriver driver;
	
	//********************************************************************************************************************
	//Function Description 	: To Initialize the execution
	//Author				: Antony Ignatius Thenraja F
	//Last Modified			: 28-Feb-2021
	//********************************************************************************************************************
	
	public void InitScript() {
		
		try {
			
			LinkedHashMap<String, String> tempTestData = null;

			tagName = configReader.getPropertyValue("tagToBeExecuted");
			sheetName = configReader.getPropertyValue("sheetName");
			
			printLog("INFO", "INFO : Tag To be executed = " + tagName);
			printLog("INFO", "INFO : Sheet Name = " + sheetName);
			
			printLog("INFO", "INFO : Screenshots Initialization for the TC " + tagName);
			screenshotInit(tagName);
			
			printLog("INFO", "INFO : Extent Report Initialization for the TC " + tagName);
			ExtentReportClass.initializeExtentReport();
			currentlyExecutedTest = ExtentReportClass.InitializeExecutionReportUsingExtent(tagName);
			
			tempTestData = loadTestData();
			
			printLog("INFO", "INFO : Test Data = " + tempTestData.toString());
			
		}
		catch (Exception e) {
			
		}
		
	}
	
	//********************************************************************************************************************
		//Function Description 	: To load the test data from excel
		//Author				: Antony Ignatius Thenraja F
		//Last Modified			: 28-Feb-2021
		//********************************************************************************************************************
	
	public LinkedHashMap<String, String> loadTestData() {
		
		LinkedHashMap<String, String> tempTestData = null;
		
		try {
			sheetName = configReader.getPropertyValue("sheetName");
			tempTestData = excel.loadExcelData(sheetName, tagName);
		}
		catch(Exception e) {
			printLog("INFO", "INFO : Error in loaded Excel");
		}
		
		return tempTestData;
		
	}
	
	//********************************************************************************************************************
	//Function Description 	: To Initialize the the screen capture path
	//Author				: Antony Ignatius Thenraja F
	//Last Modified			: 28-Feb-2021
	//********************************************************************************************************************
	
	public void screenshotInit(String testCaseName) {
		
		try {

			String screenCurrentImagePath = "";
			screenCurrentImagePath = configReader.getPropertyValue("ImageFolderPath") + testCaseName;
			
			File theDir = new File(screenCurrentImagePath);
			
			if(theDir.exists()) {
				printLog("INFO", "INFO : File already exist : " + theDir.getName());
			}
			
			printLog("INFO", "INFO : Creating directory to store the screenshot: " + theDir.getName());
			
			boolean result = false;
			
			try {
				theDir.mkdir();
				result = true;
			}
			catch (Exception s) {
				s.printStackTrace();
				printLog("FAIL", "FAIL : Error in creating the directory");
			}
			
			if(result) {
				printLog("INFO", "INFO : DIR created successfully " + theDir.getName());
			}
			else {
				printLog("INFO", "INFO : DIR NOT created " + theDir.getName());
			}
			
		} catch (Exception e) {
			e.printStackTrace();
			printLog("FAIL", "FAIL : Error in screenshotInit");
		}
		
	}
	
	//********************************************************************************************************************
	//Function Description 	: To enable logging
	//Author				: Antony Ignatius Thenraja F
	//Last Modified			: 28-Feb-2021
	//********************************************************************************************************************
	
	public void printLog(String logType, String logDescription) {
		
		switch (logType.toLowerCase()) {
		
		case "info":
			logger.info("");
			logger.info("[Thread ID - " + threadId + " ] ----> " + logDescription);
			break;
			
		case "debug":
			logger.info("");
			logger.debug("[Thread ID - " + threadId + " ] ----> " + logDescription);
			break;
			
		case "warn":
			logger.info("");
			logger.warn("[Thread ID - " + threadId + " ] ----> " + logDescription);
			break;
			
		case "fail":
			logger.info("");
			System.out.println("[Thread ID - " + threadId + " ] ----> " + logDescription);
			Assert.fail(logDescription);
			break;
		
		}
		
	}
	
}
