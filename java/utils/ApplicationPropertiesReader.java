package utils;

import java.io.BufferedReader;
import java.io.FileReader;
import java.util.Properties;

public class ApplicationPropertiesReader {
	
	public Properties appProperties;
	public final String appPropertiesFilePath = "projectResources//application.properties";
	
	BufferedReader appReader;
	
	public ApplicationPropertiesReader() {
		
		try {
			
			appReader = new BufferedReader(new FileReader(appPropertiesFilePath));
			appProperties = new Properties();
			
			appProperties.load(appReader);
			appReader.close();
			
		}
		catch (Exception e) {
			e.printStackTrace();
			//throw new Exception ("Error in Loading application.properties file from " + appPropertiesFilePath);
		}
		
	}
	
	public String getPropertyValue(String propertyValue) throws Exception {
		
		String value = "";
		
		try {
			
			appProperties.load(new FileReader(appPropertiesFilePath));
			value = appProperties.getProperty(propertyValue);
			
		}
		catch (Exception e) {
			e.printStackTrace();
			throw new Exception ("Error in getting value for " + propertyValue + " from application.properties file at " + appPropertiesFilePath);
		}
		
		return value;
		
	}
	
	public String getDriverPathForChrome() throws Exception {
		
		try {
			
			String driverPath = appProperties.getProperty("ChromeDriverPath");
			if (driverPath != null) return driverPath;
			else throw new Exception ("Chromedriver path NOT specified in application.properties file");
			
		}
		catch (Exception e) {
			e.printStackTrace();
			throw new Exception ("Error in getting chromedriver path from application.properties file at " + appPropertiesFilePath);
		}
		
	}

}
