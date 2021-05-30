package pages;

import utils.ApplicationPropertiesReader;
import utils.InitScriptClass;

public class GooglePage extends InitScriptClass {
	
	ApplicationPropertiesReader appReader = new ApplicationPropertiesReader();
	
	public void launchGoogle() {
		webDriver.get("https://www.google.com/");
		waitFor(10);
		takeScreenshot("Google Landing Page", "GoogleSearch", "hi");
	}
	
	public void searchInGoogle() {
		
	}

}
