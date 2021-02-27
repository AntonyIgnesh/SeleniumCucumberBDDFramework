package utils;

import java.io.BufferedReader;
import java.io.FileReader;
import java.util.Properties;

public class ConfigReader {
	
	public Properties configProperties;
	public final String configPropertiesFilePath = "projectResources//config.properties";
	
	BufferedReader configReader;
	
	public ConfigReader() {

		try {
			
			configReader = new BufferedReader(new FileReader(configPropertiesFilePath));
			configProperties = new Properties();
			
			configProperties.load(configReader);
			configReader.close();
			
		}
		catch (Exception e) {
			e.printStackTrace();
			//throw new Exception ("Error in Loading config.properties file from " + configPropertiesFilePath);
		}
		
	}
	
	public String getPropertyValue(String propertyValue) throws Exception {
		
		String value = "";
		
		try {
			
			configProperties.load(new FileReader(configPropertiesFilePath));
			value = configProperties.getProperty(propertyValue);
			
		}
		catch (Exception e) {
			e.printStackTrace();
			//throw new Exception ("Error in getting value for " + propertyValue + " from application.properties file at " + configPropertiesFilePath);
		}
		
		return value;
		
	}

}
