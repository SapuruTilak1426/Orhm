package OrangeHrmApplicationTestDataFiles;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Properties;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.testng.annotations.Test;

import BaseTest.BaseTest;
import Utility.Log;

public class Employeelist1 extends BaseTest
{
	FileInputStream OrangeHrmApplicationpropertiesFile;
	Properties Properties;

	@Test(priority = 1, description = " Validating OrangeHRM Application Employee list")
	public void employeeListValidation() throws IOException {
		FileInputStream employeeListValidation = new FileInputStream(
				"./src/main/java/OrangeHrmApplicationTestDataFiles/EmployeeList.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(employeeListValidation);
		XSSFSheet EmployeeList = workbook.getSheet("EmployeeList");
		OrangeHrmApplicationpropertiesFile = new FileInputStream(
				"./src/main/java/Config/OrangeHrmApplication.Properties");
		Properties = new Properties();
		Properties.load(OrangeHrmApplicationpropertiesFile);

		Row row = EmployeeList.getRow(1);

		Cell expectedLoginPanelRowofCell = row.getCell(0);

		String ExpctedLoginPanelText = expectedLoginPanelRowofCell.getStringCellValue();
		Log.info("This Expected Login Panel Text Is:-" + ExpctedLoginPanelText);

		By loginPanelproperty = By.id(Properties.getProperty("LogInPageLoginPanelProperty"));
		WebElement loginPanel = driver.findElement(loginPanelproperty);
		String actualLoginPanelText = loginPanel.getText();
		Log.info("The Actual Login Panel Text Is:- " + actualLoginPanelText);
		Cell Actual_LogInPageTextRowOfCell = row.createCell(1);
		Actual_LogInPageTextRowOfCell.setCellValue(actualLoginPanelText);
		Cell loginTextTestResult = row.createCell(2);

		if (actualLoginPanelText.equals(ExpctedLoginPanelText)) {
			Log.info("Successfully Navigated to OrangeHrm Login Page:-Pass");
			loginTextTestResult.setCellValue("Pass");
		} else {
			Log.info("Failed to Navigate to OrangeHrm Login Page:-Fail");
			loginTextTestResult.setCellValue("Fail");
		}
		System.out.println();

		Cell LogInPageTitle = row.getCell(3);
		String expected_LogInPageTitle = LogInPageTitle.getStringCellValue();
		Log.info("The Expected Title of the OrangeHrm Login Page is:-" + expected_LogInPageTitle);

		String actual_LoginPageTitle = driver.getTitle();
		Log.info("The Actual Title of the OrangeHrm Login Page is:-" + actual_LoginPageTitle);
		Cell actual_LoginPageTitleRowofCell = row.createCell(4);
		actual_LoginPageTitleRowofCell.setCellValue(actual_LoginPageTitle);
		Cell titleTestResult = row.createCell(5);

		if (actual_LoginPageTitle.equalsIgnoreCase(expected_LogInPageTitle)) {
			Log.info("The Title of the OrangeHrm Page is Matched:-Pass");
			titleTestResult.setCellValue("Pass");
		} else {
			Log.info("The Title of the OrangeHrm Page is Not Matched:-Fail");
			titleTestResult.setCellValue("Fail");
		}
		System.out.println();

		Cell usernameTextDataRowOfCell = row.getCell(6);
		String usernameTextData = usernameTextDataRowOfCell.getStringCellValue();
		By usernameProperty = By.id(Properties.getProperty("LogInPageUserNameProperty"));
		WebElement username = driver.findElement(usernameProperty);
		username.sendKeys(usernameTextData);

		Cell passwordTextDataRowOfCell = row.getCell(7);
		String passwordTextData = passwordTextDataRowOfCell.getStringCellValue();
		By passwordProperty = By.id(Properties.getProperty("LogInPagePasswordProperty"));
		WebElement password = driver.findElement(passwordProperty);
		password.sendKeys(passwordTextData);

		By loginButtonProperty = By.id(Properties.getProperty("LogInPageLoginButtonProperty"));
		WebElement loginButton = driver.findElement(loginButtonProperty);
		loginButton.click();

		Cell expectedHomePageTextRowofCell = row.getCell(8);
		String expectedHomePageText = expectedHomePageTextRowofCell.getStringCellValue();
		Log.info("The Expected OrangeHrm Home Page Text is:-" + expectedHomePageText);

		By welcomeAdminProperty = By.id("welcome");
		WebElement welcomeadmin = driver.findElement(welcomeAdminProperty);
		String actualHomePageText = welcomeadmin.getText();
		Cell actualHomePageTextRowOfCell = row.createCell(19);
		actualHomePageTextRowOfCell.setCellValue(actualHomePageText);
		Log.info("The Actual OrangeHrm Home Page Text is:-" + actualHomePageText);

		Cell homePageTestResult = row.createCell(10);
		if (actualHomePageText.contains(expectedHomePageText)) {
			Log.info("Successfully Navigated to OrangeHrm Homepage:-Pass");
			homePageTestResult.setCellValue("Pass");
		} else {
			Log.info("Failed to Navigate to OrangeHrm Homepage:-Fail");
			homePageTestResult.setCellValue("Fail");
		}
		System.out.println();

		By pimProperty = By.id(Properties.getProperty("HomePagePimProperty"));
		WebElement pim = driver.findElement(pimProperty);

		Actions mouseHoverOperation = new Actions(driver);
		mouseHoverOperation.moveToElement(pim).build().perform();

		By employeeListProperty = By.id(Properties.getProperty("HomePageEmployeeListProperty"));
		WebElement employeeList = driver.findElement(employeeListProperty);
		employeeList.click();

		Cell expectedEmployeeListTextRowofCell = row.getCell(11);
		String expectedEmployeeListText = expectedEmployeeListTextRowofCell.getStringCellValue();
		Log.info("The Expected OrangeHrm EmployeeList Page Text is:-" + expectedEmployeeListText);

//		/html/body/div[1]/div[3]/div[1]/div[1]/h1
		By EmployeeListTextProperty = By.xpath(Properties.getProperty("EmployeeListEmployeeListTextProperty"));
		WebElement EmployeeListText = driver.findElement(EmployeeListTextProperty);
		String actualEmployeeListText = EmployeeListText.getText();
		Log.info("The Actual  OrangeHrm EmployeeList Page Text is:- " + actualEmployeeListText);

		Cell actualEmployeeListTextRowofcell = row.createCell(12);
		actualEmployeeListTextRowofcell.setCellValue(actualEmployeeListText);

		Cell EmployeePageTestResult = row.createCell(13);
		if (actualEmployeeListText.equals(expectedEmployeeListText)) {
			Log.info("Successfully Navigated to OrangeHrm Employee List Page:-Pass");
			EmployeePageTestResult.setCellValue("Pass");
		} else {
			Log.info("Failed to Navigate to OrangeHrm Employee List Page:-Fail");
			EmployeePageTestResult.setCellValue("Fail");
		}


//		/html/body/div[1]/div[3]/div[2]/div/form/div[4]/table/tbody/tr[1]/td[2]/a
//		/html/body/div[1]/div[3]/div[2]/div/form/div[4]/table/tbody/tr[1]/td[3]/a
//		/html/body/div[1]/div[3]/div[2]/div/form/div[4]/table/tbody/tr[1]/td[4]/a
		// /html/body/div[1]/div[3]/div[2]/div/form
	    for (int page = 1; page <= 3; page++) {
	        // Create a new sheet for each page
	        XSSFSheet resultSheet = workbook.createSheet("Result_Page_" + page);

	        // Initialize row number for the result sheet
	        int resultRowNum = 0;

	        // Navigate to the next page if not the first page
	        if (page > 1) {
	            By nextPageButtonProperty = By.xpath("/html/body/div[1]/div[3]/div[2]/div/form/div[5]/ul/li[7]/a");
	            WebElement nextPageButton = driver.findElement(nextPageButtonProperty);
	            nextPageButton.click();
	        }

	        // Locate the web table on the current page
	        By webTableProperty = By.xpath(Properties.getProperty("EmployeeListEmployeeListWebTableProperty"));
	        WebElement webTable = driver.findElement(webTableProperty);

	        // Locate all rows of the web table
	        List<WebElement> webTableRows = webTable.findElements(By.tagName("tr"));

	        // Iterate through each row in the web table
	        for (int rowIndex = 0; rowIndex < webTableRows.size(); rowIndex++) {
	            // Create a new row in the result sheet
	            Row resultRow = resultSheet.createRow(resultRowNum++);

	            // Locate all cells of the current row
	            List<WebElement> rowOfCells = webTableRows.get(rowIndex).findElements(By.tagName("td"));

	            // Initialize cell index
	            int cellIndex = 0;

	            // Iterate through each cell in the row
	            for (WebElement cell : rowOfCells) {
	                // Get the text of the cell
	                String cellText = cell.getText();

	                // Write the text to the result sheet
	                resultRow.createCell(cellIndex++).setCellValue(cellText);

	                // Print the text to the console
	                System.out.print(cellText + " | ");
	            }

	            // Move to the next line in the console
	            System.out.println();
	        }
	    }


	
		FileOutputStream FileOutput = new FileOutputStream(
				"./src/main/java/OrangeHrmApplicationTestResultFiles/EmployeeListResult.xlsx");
		workbook.write(FileOutput);
		
		
		
		

	}

}



