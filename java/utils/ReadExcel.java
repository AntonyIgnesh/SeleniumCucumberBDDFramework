package utils;

import java.io.FileInputStream;
import java.util.LinkedHashMap;

import org.apache.log4j.LogManager;
import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.Assert;

public class ReadExcel {
	
	public static Logger logger = LogManager.getLogger(ReadExcel.class);
	ConfigReader configReader = new ConfigReader();
	
	public LinkedHashMap<String, String> loadExcelData(String sheetName, String tagName) {
		
		LinkedHashMap<String, String> excelValue = new LinkedHashMap<String, String>();
		
		Workbook workBook = null;
		Sheet sheet = null;
		int i,j = 0;
		boolean matchFound = false;
		
		try {
			
			String testDataPath = configReader.getPropertyValue("excelPath");
			
			logger.info("INFO : Reading Excel");
			
			FileInputStream fis = new FileInputStream(testDataPath);
			
			String fileExtensionName = testDataPath.substring(testDataPath.indexOf("."));
			
			if (fileExtensionName.equalsIgnoreCase(".xlsx")) {
				
				workBook = new XSSFWorkbook(fis);
				
			}
			else if (fileExtensionName.equalsIgnoreCase(".xls")) {
				
				workBook = new HSSFWorkbook(fis);
				
			}
			
			sheet = workBook.getSheet(sheetName);
			
			int rowCount = sheet.getLastRowNum();
			int colCount = sheet.getRow(0).getLastCellNum();
			
			logger.info("INFO : Reading Excel sheet " + sheet.getSheetName() + " with " + rowCount + " rows and " + colCount + " columns");
			
			Row excelHeader = sheet.getRow(0);
			
			for (i = 1; i <= rowCount; i++) {
				
				if (sheet.getRow(i).getCell(0).getStringCellValue().equalsIgnoreCase(tagName)) {
					
					matchFound = true;
					
					for (j = 1; j < colCount; j++) {
						
						if (sheet.getRow(i).getCell(j) != null) {
							
							try {
								
								excelValue.put(excelHeader.getCell(j).getStringCellValue(), sheet.getRow(i).getCell(j).getStringCellValue());
								
							}
							catch (Exception f) {
								f.printStackTrace();
								excelValue.put(excelHeader.getCell(j).getStringCellValue(), "");
							}
							
						}
						else {
							excelValue.put(excelHeader.getCell(j).getStringCellValue(), "");
						}
						
					}
					
					i = rowCount + 1;
					
				}
				
			}
			
			fis.close();
			
			workBook.close();
			
			if (matchFound == false) {
				
				logger.info("FAIL : Matching row NOT found in excel. Please check the excel path or sheetname");
				
				Assert.fail("FAIL : Matching row NOT found in excel. Please check the excel path or sheetname");
				
			}
			
		}
		catch(Exception e) {
			
			logger.info("FAIL : Error in loadExcelData");
			
			Assert.fail("FAIL : Error in loadExcelData");
			
		}
		
		return excelValue;
		
	}

}
