package utils;

import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTabStop;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTabJc;
import org.testng.Assert;

import com.aventstack.extentreports.ExtentTest;
import com.aventstack.extentreports.MediaEntityBuilder;
import com.aventstack.extentreports.Status;

import ru.yandex.qatools.ashot.AShot;
import ru.yandex.qatools.ashot.Screenshot;
import ru.yandex.qatools.ashot.shooting.ShootingStrategies;

import java.awt.Graphics2D;
import java.awt.Rectangle;
import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.lang.System.Logger.Level;
import java.math.BigInteger;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Base64;
import java.util.Date;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.concurrent.TimeUnit;
import java.util.logging.ConsoleHandler;
import java.util.logging.Handler;

import javax.imageio.ImageIO;

import org.apache.commons.io.FileUtils;
import org.apache.log4j.LogManager;
import org.apache.log4j.Logger;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.UnderlinePatterns;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;

public class InitScriptClass {
	
	ReadExcel excel = new ReadExcel();
	
	public static Logger logger = LogManager.getLogger(InitScriptClass.class);
	
	public long threadId = Thread.currentThread().getId();
	
	ApplicationPropertiesReader appReader = new ApplicationPropertiesReader();
	
	ConfigReader configReader = new ConfigReader();
	
	public String tagName = "";
	public String sheetName = "";
	public String browserName = "";
	public String extentReportImageOption = "";
	
	public static String screenshotName;
	
	public static ExtentTest currentlyExecutedTest;
	
	public static WebDriver webDriver;
	
	public static LinkedHashMap<String, String> pathToImage = new LinkedHashMap<String, String>();
	
	/****************************************************************************************************************************************
	 * 
	 * Function Description 	: To Initialize the execution
	 * Author					: Antony Ignatius Thenraja F
	 * Modified By				:
	 * Last Modified			: 28-Feb-2021
	 * 
	 * **************************************************************************************************************************************/
	
	public LinkedHashMap<String, String> InitScript() {

		LinkedHashMap<String, String> tempTestData = null;

		try {
			
			tagName = configReader.getPropertyValue("tagToBeExecuted");
			sheetName = configReader.getPropertyValue("sheetName");
			browserName = configReader.getPropertyValue("Browser");
			extentReportImageOption = configReader.getPropertyValue("ExtentReportImageOption");
			
			printLog("INFO", "INFO : Tag To be executed = " + tagName);
			printLog("INFO", "INFO : Sheet Name = " + sheetName);
			printLog("INFO", "INFO : Broswer Name = " + browserName);
			printLog("INFO", "INFO : Extent Report Image Option = " + extentReportImageOption);
			
			printLog("INFO", "INFO : Screenshots Initialization for the TC " + tagName);
			screenshotInit(tagName);
			
			printLog("INFO", "INFO : Extent Report Initialization for the TC " + tagName);
			ExtentReportClass.initializeExtentReport();
			currentlyExecutedTest = ExtentReportClass.InitializeExecutionReportUsingExtent(tagName);
			
			tempTestData = loadTestData();
			
			printLog("INFO", "INFO : Test Data = " + tempTestData.toString());
			
			printLog("INFO", "INFO : Browser Initialization for the TC " + tagName);
			
			switch (browserName.toLowerCase()) {
			
			case "chrome":
				launchChromeDriver();
				break;
				
			case "IE":
				launchIEDriver();
				break;
				
			}
			
			@SuppressWarnings("rawtypes")
			Iterator it = tempTestData.entrySet().iterator();
			
			while (it.hasNext()) {
				
				@SuppressWarnings("rawtypes")
				Map.Entry pair = (Map.Entry) it.next();
				
				printLog("INFO", "INFO : | " + pair.getKey() + " | " + pair.getValue() + " |");
				
				it.remove();
				
			}
			
		}
		catch (Exception e) {
			printLog("FAIL", "FAIL : Error in InitScript");
		}

		return tempTestData;
		
	}
	
	/****************************************************************************************************************************************
	 * 
	 * Function Description 	: To load the test data from excel
	 * Author					: Antony Ignatius Thenraja F
	 * Modified By				:
	 * Last Modified			: 28-Feb-2021
	 * 
	 * **************************************************************************************************************************************/
	
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
	
	/****************************************************************************************************************************************
	 * 
	 * Function Description 	: To Initialize the the screen capture path
	 * Author					: Antony Ignatius Thenraja F
	 * Modified By				:
	 * Last Modified			: 28-Feb-2021
	 * 
	 * **************************************************************************************************************************************/
	
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

	/****************************************************************************************************************************************
	 * 
	 * Function Description 	: To launch the IE browser
	 * Author					: Antony Ignatius Thenraja F
	 * Modified By				:
	 * Last Modified			: 28-Feb-2021
	 * 
	 * **************************************************************************************************************************************/
	
	@SuppressWarnings("deprecation")
	public void launchIEDriver() {
		
		
		try {
			
			closeIEDriver();
			
			printLog("INFO", "INFO : IEDriver closed");
			
			DesiredCapabilities cap = DesiredCapabilities.internetExplorer();
			
			System.setProperty("webdriver.chrome.driver", System.getProperty("user.dir") + "/drivers/IEDriverServer");
			
			printLog("INFO", "INFO : IE Driver launch Started");
			
			printLog("INFO", "INFO : Adding Capabilities for IE Driver");
			
			cap.setCapability("ignoreProtectedModeSettings", true);
			cap.setCapability(InternetExplorerDriver.ENABLE_ELEMENT_CACHE_CLEANUP, true);
			cap.setCapability("ignoreZoomSetting", true);
			
			printLog("INFO", "INFO : Capabilities set for IE Driver");
			
			webDriver = new InternetExplorerDriver(cap);
			
			printLog("INFO", "INFO : IE Driver launched successfully");
			
		}
		catch (Exception e) {
			
		}
		
	}

	/****************************************************************************************************************************************
	 * 
	 * Function Description 	: To quit the IE browser
	 * Author					: Antony Ignatius Thenraja F
	 * Modified By				:
	 * Last Modified			: 28-Feb-2021
	 * 
	 * **************************************************************************************************************************************/
	
	public void closeIEDriver() {
		
		try {
			
			printLog("INFO", "INFO : Killing IEDriver");
			
			Runtime.getRuntime().exec("taskkill /F /IM IEDriverServer.exe /T");
			
			printLog("INFO", "INFO : Killed IEDriver");
			
			printLog("INFO", "INFO : Killing iexplore.exe");
			
			Runtime.getRuntime().exec("taskkill /F /IM iexplore.exe /T");
			
			printLog("INFO", "INFO : Killed iexplore.exe");
			
			Runtime.getRuntime().exec("RunDLL32.exe InetCpl.cpl,ClearMyTracksByProcess 'pid get 0'");
			
		}
		catch (Exception e) {
			e.printStackTrace();
			printLog("FAIL", "FAIL : Error in closeIEDriver");
		}
		
	}

	/****************************************************************************************************************************************
	 * 
	 * Function Description 	: To launch the Chrome Browser
	 * Author					: Antony Ignatius Thenraja F
	 * Modified By				:
	 * Last Modified			: 28-Feb-2021
	 * 
	 * **************************************************************************************************************************************/
	
	public void launchChromeDriver() {
		
		try {
			
			printLog("INFO", "INFO : Chrome Driver launch Started");

			System.setProperty("webdriver.chrome.driver", System.getProperty("user.dir") + "/drivers/chromedriver");
			
			ChromeOptions option = new ChromeOptions();
			
			option.setExperimentalOption("useAutomationExtension", false);
			
			option.addArguments("--start-maximized");
			
			webDriver = new ChromeDriver(option);
			
			webDriver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
			
			webDriver.manage().window().maximize();
			
			printLog("INFO", "INFO : Chrome Driver launched successfully");
			
		}
		catch (Exception e) {
			e.printStackTrace();
			printLog("FAIL", "FAIL : Error in launchChromeDriver");
		}
		
	}

	/****************************************************************************************************************************************
	 * 
	 * Function Description 	: To quit the Chrome browser
	 * Author					: Antony Ignatius Thenraja F
	 * Modified By				:
	 * Last Modified			: 28-Feb-2021
	 * 
	 * **************************************************************************************************************************************/
	
	public void closeChromeDriver() {
		
		try {
			
			printLog("INFO", "INFO : Killing ChromeDriver");
			
			Runtime.getRuntime().exec("taskkill /F /IM chromedriver.exe /T");
			
			printLog("INFO", "INFO : Killed ChromeDriver");
			
			waitFor(2);
			
		}
		catch (Exception e) {
			e.printStackTrace();
			printLog("FAIL", "FAIL : Error in closeChromeDriver");
		}
		
	}

	/****************************************************************************************************************************************
	 * 
	 * Function Description 	: To wait for seconds
	 * Author					: Antony Ignatius Thenraja F
	 * Modified By				:
	 * Last Modified			: 28-Feb-2021
	 * 
	 * **************************************************************************************************************************************/
	
	public void waitFor (long secondsToWait) {
		
		try {
			
			printLog("INFO", "INFO : Thread will wait for " + secondsToWait + " seconds");
			
			Thread.sleep(secondsToWait * 1000);
			
			printLog("INFO", "INFO : Thread waited for " + secondsToWait + " seconds");
			
		}
		catch (Exception e) {
			e.printStackTrace();
			printLog("WARN", "WARN : Error in waitFor");
		}
		
	}

	/****************************************************************************************************************************************
	 * 
	 * Function Description 	: To update the extent reports with the screenshots
	 * Author					: Antony Ignatius Thenraja F
	 * Modified By				:
	 * Last Modified			: 28-Feb-2021
	 * 
	 * **************************************************************************************************************************************/
	
	public void updateExtentReport(String status, String message, String fullPage) {
		
		try {
			
			String ssPath = "";
			extentReportImageOption = configReader.getPropertyValue("ExtentReportImageOption");
			
			printLog("INFO", "INFO : Updating Extent report for " + message + " with " + status + " status and full page option as " + fullPage);
			
			switch (extentReportImageOption.toUpperCase()) {
			
			case "BASE64":
				
				switch (fullPage.toUpperCase()) {
				
				case "YES":
					
					ssPath = takeBase64Screenshot(screenshotName, tagName, message);
					
					break;
					
				case "NO":
					
					ssPath = takeBase64SingleScreenshot(screenshotName, tagName, message);
					
					break;
				
				}

				if (status.equalsIgnoreCase("pass")) {
					
					currentlyExecutedTest.log(Status.PASS, message, MediaEntityBuilder.createScreenCaptureFromBase64String(ssPath).build());
					
				}
				else {
					
					currentlyExecutedTest.log(Status.FAIL, message, MediaEntityBuilder.createScreenCaptureFromBase64String(ssPath).build());
					
				}
				
				break;
				
			case "LOCAL":
				
				switch (fullPage.toUpperCase()) {
				
				case "YES":
					
					ssPath = takeScreenshot(screenshotName, tagName, message);
					
					break;
					
				case "NO":
					
					ssPath = takeSingleScreenshot(screenshotName, tagName, message);
					
					break;
				
				}
				
				if (status.equalsIgnoreCase("pass")) {
					
					currentlyExecutedTest.log(Status.PASS, message, MediaEntityBuilder.createScreenCaptureFromPath(ssPath).build());
					
				}
				else {
					
					currentlyExecutedTest.log(Status.FAIL, message, MediaEntityBuilder.createScreenCaptureFromPath(ssPath).build());
					
				}
				
				break;
				
			}
			
			printLog("INFO", "INFO : Updated Extent report for " + message + " with " + status + " status and full page option as " + fullPage);
			
		}
		catch (Exception e) {
			e.printStackTrace();
			printLog("FAIL", "FAIL : Error in updateExtentReport for " + message);
		}
		
	}

	/****************************************************************************************************************************************
	 * 
	 * Function Description 	: To update the extent report as PASS
	 * Author					: Antony Ignatius Thenraja F
	 * Modified By				:
	 * Last Modified			: 28-Feb-2021
	 * 
	 * **************************************************************************************************************************************/
	
	public void extentReportPass(String message) {
		
		try {
			
			updateExtentReport("Pass", message, "Yes");
			
		}
		catch (Exception e) {
			e.printStackTrace();
			printLog("FAIL", "FAIL : Error in extentReportPass");
		}
		
	}

	/****************************************************************************************************************************************
	 * 
	 * Function Description 	: To update extent report with single screenshot
	 * Author					: Antony Ignatius Thenraja F
	 * Modified By				:
	 * Last Modified			: 28-Feb-2021
	 * 
	 * **************************************************************************************************************************************/
	
	public void extentReportPassSingle(String message) {
		
		try {
			
			updateExtentReport("Pass", message, "No");
			
		}
		catch (Exception e) {
			e.printStackTrace();
			printLog("FAIL", "FAIL : Error in extentReportPassSingle");
		}
		
	}	

	/****************************************************************************************************************************************
	 * 
	 * Function Description 	: To take screenshot
	 * Author					: Antony Ignatius Thenraja F
	 * Modified By				:
	 * Last Modified			: 28-Feb-2021
	 * 
	 * **************************************************************************************************************************************/
	
	public String getScreenShotImage(String screenshotName, String testCaseName, String message, boolean fullScreenshot) {
		
		String screenCurrentImagePath = "";
		String path = "";
		Date date = new Date();
		SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MM-yy'_'HH-mm-ss");
		
		try {

			screenCurrentImagePath = configReader.getPropertyValue("ImageFolderPath") + testCaseName;
			
			path = "\\" + screenshotName + "_" + dateFormat.format(date) + ".png";
			BufferedImage image = null;
			List<String> result = new ArrayList<>();
			String ssOption = configReader.getPropertyValue("Browser");
			
			switch (ssOption.toLowerCase()) {
			
			case "chrome":
				
				if (fullScreenshot==false) {
					
					File srcFile = ((TakesScreenshot) webDriver).getScreenshotAs(OutputType.FILE);
					FileUtils.copyFile(srcFile, new File(screenCurrentImagePath + path));
					
					printLog("INFO", "INFO : Took single Screenshot in Chrome");
					
				}
				else {
					
					Screenshot screenshot = null;
					
					screenshot = new AShot().shootingStrategy(ShootingStrategies.viewportPasting(1000)).takeScreenshot(webDriver);
					
					ImageIO.write(screenshot.getImage(), "png", new File(screenCurrentImagePath + path));
					
					printLog("INFO", "INFO : Took full Screenshot in Chrome");
					
				}
				
				break;
				
			case "ie":
				
				Rectangle allScreenBounds = new Rectangle();
				
				allScreenBounds.width = 1921;
				allScreenBounds.height = 1077;
				
				if (fullScreenshot==false) {
					
					image = new Robot().createScreenCapture(new Rectangle(allScreenBounds));
					
					ImageIO.write(image, "png", new File(screenCurrentImagePath + path));
					
					printLog("INFO", "INFO : Took single Screenshot in IE");
					
				}
				else {
					
					Robot robot = new Robot();
					
					JavascriptExecutor jsExec = (JavascriptExecutor) webDriver;
					
					jsExec.executeScript("window.scrollTo(0, 0);");
					
					Long innerHeight = (Long) jsExec.executeScript("return window.innerHeight;");
					
					Long scroll = innerHeight;
					
					Long scrollHeight = (Long) jsExec.executeScript("return document.body.scrollHeight;");
					
					scrollHeight = scrollHeight + scroll + scroll;
					
					int counter = 1;
					
					do {
						
						image = new Robot().createScreenCapture(new Rectangle(allScreenBounds));
						
						ImageIO.write(image, "png", new File(screenCurrentImagePath + "\\" + screenshotName + "_" + dateFormat.format(date) + "_" + counter + ".png"));
						
						result.add(screenCurrentImagePath + "\\" + screenshotName + "_" + dateFormat.format(date) + "_" + counter + ".png");
						
						robot.keyPress(KeyEvent.VK_PAGE_DOWN);
						
						counter = counter + 1;
						
						Thread.sleep(1000);
						
						innerHeight = innerHeight + scroll;
						
					} while (scrollHeight >= innerHeight);
					
					int imageTotalHeight = 0;
					int imageWidth = 0;
					
					BufferedImage imageList;
					
					for (String s: result) {
						
						imageList = ImageIO.read(new File(s));
						imageTotalHeight += imageList.getHeight();
						imageWidth = imageList.getWidth();
						
					}
					
					int heightCurr = 0;
					BufferedImage finalImage = new BufferedImage(imageWidth, imageTotalHeight, BufferedImage.TYPE_INT_RGB);
					Graphics2D g2d = finalImage.createGraphics();
					
					for (String s: result) {
						imageList = ImageIO.read(new File(s));
						g2d.drawImage(imageList, 0, heightCurr, null);
						heightCurr += imageList.getHeight();
					}
					g2d.dispose();
					
					ImageIO.write(finalImage, "png", new File(screenCurrentImagePath + path));
					
					printLog("INFO", "INFO : Took full Screenshot in IE");
					
				}
				
				break;
			
			}
			
		}
		catch(Exception e) {
			e.printStackTrace();
			printLog("FAIL", "FAIL : Error in getScreenShotImage");
		}
		
		if (screenCurrentImagePath + path != null && screenCurrentImagePath + path != "") {
			pathToImage.put(message + " @ " + dateFormat.format(date), screenCurrentImagePath + path);
		}
		
		return screenCurrentImagePath + path;
		
	}

	/****************************************************************************************************************************************
	 * 
	 * Function Description 	: To take Base 64 screenshot 
	 * Author					: Antony Ignatius Thenraja F
	 * Modified By				:
	 * Last Modified			: 28-Feb-2021
	 * 
	 * **************************************************************************************************************************************/
	
	public String takeBase64Screenshot (String screenshotName, String testCaseName, String message) {
		
		File file = null;
		String base64Image = null;
		
		try {
			
			file = new File(getScreenShotImage(screenshotName, testCaseName, message, true));
			
			@SuppressWarnings("resource")
			FileInputStream imageInFile = new FileInputStream(file);
			byte imageData[] = new byte[(int) file.length()];
			imageInFile.read(imageData);
			base64Image = Base64.getEncoder().encodeToString(imageData);
			
		}
		catch (Exception e) {
			e.printStackTrace();
			printLog("FAIL", "FAIL : Error in takeBase64Screenshot");
		}

		return base64Image;
		
	}
	
	public String takeBase64SingleScreenshot (String screenshotName, String testCaseName, String message) {
		
		File file = null;
		String base64Image = null;
		
		try {
			
			file = new File(getScreenShotImage(screenshotName, testCaseName, message, false));
			
			@SuppressWarnings("resource")
			FileInputStream imageInFile = new FileInputStream(file);
			byte imageData[] = new byte[(int) file.length()];
			imageInFile.read(imageData);
			base64Image = Base64.getEncoder().encodeToString(imageData);
			
		}
		catch (Exception e) {
			e.printStackTrace();
			printLog("FAIL", "FAIL : Error in takeBase64SingleScreenshot");
		}

		return base64Image;
		
	}
	
	public String takeScreenshot (String screenshotName, String testCaseName, String message) {
		
		String localImage = null;
		
		try {
			
			localImage = getScreenShotImage(screenshotName, testCaseName, message, true);
			
		}
		catch(Exception e) {
			e.printStackTrace();
			printLog("FAIL", "FAIL : Error in takeScreenshot");
		}
		
		return localImage;
		
	}
	
	public String takeSingleScreenshot (String screenshotName, String testCaseName, String message) {
		
		String localImage = null;
		
		try {
			
			localImage = getScreenShotImage(screenshotName, testCaseName, message, false);
			
		}
		catch(Exception e) {
			e.printStackTrace();
			printLog("FAIL", "FAIL : Error in takeSingleScreenshot");
		}
		
		return localImage;
		
	}
	
	public void generateWordReport(String status) {
		
		String screenCurrentImagePath = "";
		Date date = new Date();
		SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MM-yy'_'HH-mm-ss");
		String path = "";
		
		try {
			
			tagName = configReader.getPropertyValue("tagToBeExecuted");
			screenCurrentImagePath = configReader.getPropertyValue("WordReportPath");
			
			path = tagName + "_" + status + "_" + dateFormat.format(date) + ".docx";
			
			printLog("INFO", "INFO : Word Report generation started in path = " + screenCurrentImagePath + path);
			
			File fileToCopy = new File("ExecutionReportExtent.html");
			File newFile = new File("ExtentReport_" + tagName + "_" + dateFormat.format(date) + ".html");
			
			FileUtils.copyFile(fileToCopy, newFile);
			
			printLog("INFO", "INFO : Extent Report created successfully in path = " + newFile.getAbsolutePath());
			
			XWPFDocument docx = new XWPFDocument();
			XWPFParagraph headingPara = docx.createParagraph();
			XWPFRun run = headingPara.createRun();
			
			CTSectPr sectPr = docx.getDocument().getBody().addNewSectPr();
			
			XWPFHeaderFooterPolicy headerFooterPolicy = new XWPFHeaderFooterPolicy(docx, sectPr);
			XWPFHeader header = headerFooterPolicy.createHeader(XWPFHeaderFooterPolicy.DEFAULT);
			@SuppressWarnings("unused")
			XWPFFooter footer = headerFooterPolicy.createFooter(XWPFHeaderFooterPolicy.DEFAULT);
			headingPara = header.createParagraph();
			headingPara.setAlignment(ParagraphAlignment.LEFT);
			
			CTTabStop tabStop = headingPara.getCTP().getPPr().addNewTabs().addNewTab();
			tabStop.setVal(STTabJc.RIGHT);
			int twipsPerInch = 1440;
			tabStop.setPos(BigInteger.valueOf(6 * twipsPerInch));
			
			XWPFParagraph heading = docx.createParagraph();
			heading.setAlignment(ParagraphAlignment.CENTER);
			XWPFRun headingRun = heading.createRun();
			headingRun.addBreak();
			headingRun.setBold(true);
			headingRun.setColor("0071C1");
			headingRun.setUnderline(UnderlinePatterns.THICK);
			headingRun.setText("Automation Execution Report");
			headingRun.addBreak();
			
			int sCount = pathToImage.size();
			XWPFTable table = docx.createTable(sCount * 4, 1);
			
			int row = 0;
			XWPFParagraph p1;
			XWPFRun r1;
			XWPFParagraph p2;
			XWPFRun r2;
			
			for (String stepDesc: pathToImage.keySet()) {
				
				table.getRow(row).getCell(0).setColor("85C1E9");
				p1 = table.getRow(row).getCell(0).getParagraphs().get(0);
				p1.setAlignment(ParagraphAlignment.CENTER);
				r1 = p1.createRun();
				r1.setBold(true);
				r1.setText("Step Comment : " + stepDesc);
				
				p2 = table.getRow(row + 1).getCell(0).getParagraphs().get(0);
				p2.setAlignment(ParagraphAlignment.CENTER);
				r2 = p2.createRun();
				
				try {
					
					BufferedImage imageList = ImageIO.read(new File(pathToImage.get(stepDesc)));
					
					if (imageList.getHeight() > 2000) {
						r2.addPicture(new FileInputStream(pathToImage.get(stepDesc)), XWPFDocument.PICTURE_TYPE_PNG, pathToImage.get(stepDesc), Units.toEMU(520), Units.toEMU(600));
					}
					else {
						r2.addPicture(new FileInputStream(pathToImage.get(stepDesc)), XWPFDocument.PICTURE_TYPE_PNG, pathToImage.get(stepDesc), Units.toEMU(450), Units.toEMU(300));
					}
					
					row = row + 3;
					
				}
				catch (Exception e) {
					
				}
				
				printLog("INFO", "INFO : Images and comments added to word document " + screenCurrentImagePath + path);
				
				docx.write(new FileOutputStream(screenCurrentImagePath + path));
				
				printLog("INFO", "INFO : Word report created successfully in path = " + screenCurrentImagePath + path);
				
			}
			
		}
		catch(Exception e) {
			e.printStackTrace();
			printLog("FAIL", "FAIL : Error in generateWordReport");
		}
		
	}
	
	public boolean CloseExecutionForCurrentTestCase(String Msg) {
		
		try {
			
			String ssPath = "";
			
			String ssPath2 = "";
			
			if (!Msg.equalsIgnoreCase("CloseExecution")) {
				
				printLog("INFO", "INFO : CloseExecutionForCurrentTestCase Started due to failure in function");
				
				updateExtentReport("Fail", Msg, "Yes");
				
				ssPath = takeBase64Screenshot(screenshotName, tagName, Msg);
				
				printLog("INFO", "INFO : Extent report is being closed");
				
				ExtentReportClass.CleanUpExecutionReportUsingExtent(currentlyExecutedTest);
				
				printLog("INFO", "INFO : Extent report closed successfully");
				
				generateWordReport("FAIL");
				
				Assert.fail(Msg);
				
			}
			else {
				
				ssPath = takeBase64Screenshot(screenshotName, tagName, Msg);
				
				ExtentReportClass.CleanUpExecutionReportUsingExtent(currentlyExecutedTest);
				
				printLog("INFO", "INFO : Extent report closed successfully");
				
				generateWordReport("PASS");
				
			}
			
			return true;
			
		}
		catch (Exception e) {
			
			ExtentReportClass.CleanUpExecutionReportUsingExtent(currentlyExecutedTest);
			
			return true;
			
		}
		
	}
	
	/****************************************************************************************************************************************
	 * 
	 * Function Description 	: To enable logging
	 * Author					: Antony Ignatius Thenraja F
	 * Modified By				:
	 * Last Modified			: 28-Feb-2021
	 * 
	 * **************************************************************************************************************************************/
	
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
