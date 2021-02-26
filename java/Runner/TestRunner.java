package Runner;

import org.testng.annotations.Test;

//import org.junit.runner.RunWith;
import cucumber.api.CucumberOptions;
import cucumber.api.testng.AbstractTestNGCucumberTests;
//import cucumber.api.junit.Cucumber;

	@CucumberOptions(
			features = {"resources/featureFiles"}, //the path of the feature files
			glue={"stepDefinitions"}, //the path of the step definition files
			//format= {"pretty","html:test-outout", "json:json_output/cucumber.json", "junit:junit_xml/cucumber.xml"}, //to generate different types of reporting
			monochrome = true, //display the console output in a proper readable format
			strict = true, //it will check if any step is not defined in step definition file
			dryRun = false, //to check the mapping is proper between feature file and step def file
			tags = {"@GoogleSearchTest"}			
			)
	
	@Test
	public class TestRunner extends AbstractTestNGCucumberTests{
	 
	}
	
	//ORed : tags = {"@SmokeTest , @RegressionTest"} -- execute all tests tagged as @SmokeTest OR @RegressionTest
	//ANDed : tags = tags = {"@SmokeTest" , "@RegressionTest"} -- execute all tests tagged as @SmokeTest AND @RegressionTest
	

