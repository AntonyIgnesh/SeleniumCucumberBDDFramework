package utils;

import org.openqa.selenium.WebDriver;
import org.testng.Assert;
import org.apache.log4j.LogManager;
import org.apache.log4j.Logger;

public class InitScriptClass {
	
	public static Logger logger = LogManager.getLogger(InitScriptClass.class);
	
	public long threadId = Thread.currentThread().getId();
	
	ApplicationPropertiesReader appReader = new ApplicationPropertiesReader();
	
	ConfigReader configReader = new ConfigReader();
	
	WebDriver driver;
	
	public void InitScript() {
		
		try {
			
		}
		catch (Exception e) {
			
		}
		
	}
	
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
