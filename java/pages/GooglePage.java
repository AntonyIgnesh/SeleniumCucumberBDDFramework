package pages;

import utils.ApplicationPropertiesReader;

public class GooglePage {
	
	ApplicationPropertiesReader appReader = new ApplicationPropertiesReader();
	
	public void appPropTest() {
		try {
			System.out.println(appReader.getPropertyValue("Browser"));
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	public void searchInGoogle() {
		
	}

}
