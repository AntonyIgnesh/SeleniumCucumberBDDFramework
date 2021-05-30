package stepDefinitions;

import cucumber.api.java.en.Given;
import cucumber.api.java.en.Then;
import pages.GooglePage;
import pages.TestPages;
import utils.InitScriptClass;

public class TestStepDefinitions extends InitScriptClass {
	
	@Given("^Preloaded with all the setup")
	public void perconditioning () {
		System.out.println("Per-Conditioning");
	}
	
	@Then("^Login")
	public void login() {
		System.out.println("Login");
		System.out.println("Checking Report");
		new TestPages().testExtentReport();
	}
	
	@Given("^Launch Browser and go to URL")
	public void launchBrowser() {
		InitScript();
		new GooglePage().launchGoogle();
	}
	
	@Then("^Search for given data")
	public void search() {
		new TestPages().testExtentReport();
	}

}
