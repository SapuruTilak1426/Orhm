package OrangeHRMApplciationTestCases;

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
import org.openqa.selenium.Dimension;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.testng.annotations.Test;

import BaseTest.BaseTest;
import Utility.Log;

public class OragneHRMApplicatioEmployeeList extends BaseTest {
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

		By webTableProperty = By.xpath(Properties.getProperty("EmployeeListEmployeeListWebTableProperty"));
		WebElement webTable = driver.findElement(webTableProperty);

		By rowProperty = By.tagName("tr");
		List<WebElement> webTableRows = webTable.findElements(rowProperty);
		

		// Iterate through each row in the web table
		for (int rowIndex = 0; rowIndex < webTableRows.size(); rowIndex++) {
			// Get the corresponding row in the Excel sheet
			Row excelRow = EmployeeList.getRow(rowIndex);

			// going to a particular Row
			WebElement row1 = webTableRows.get(rowIndex);

			// going to a particular Row and identify the number of its Corresponding Cells
			By rowofCellProperty = By.tagName("td");
			List<WebElement> rowOfCells = row1.findElements(rowofCellProperty);

			// Initialize cell index to start from column 14
			int cellIndex = 14;

			// going to every Row of its Corresponding Cells
			for (int rowOfCellIndex = 1; rowOfCellIndex < rowOfCells.size(); rowOfCellIndex++) {
				// going to a Particular row of its Corresponding Cell
				WebElement rowOfCell = rowOfCells.get(rowOfCellIndex);

				if (excelRow == null) {
					excelRow = EmployeeList.createRow(rowIndex);
				}

				// Get the corresponding cell in the Excel sheet
				Cell excelCell = excelRow.createCell(cellIndex++);

				// get the WebTable Data from the Row of Cell
				String webTableDataText = rowOfCell.getText();

				// Print the text of the current cell
				System.out.print(webTableDataText + " | ");

				// Write the text of the current cell to the Excel cell
				excelCell.setCellValue(webTableDataText);
			}
			System.out.println();

		}
		
		 XSSFSheet Page2=workbook.createSheet("Sheet1");
		
		By NextpageButtonProperty=By.xpath("/html/body/div[1]/div[3]/div[2]/div/form/div[5]/ul/li[7]/a");
		WebElement NextpageButton=driver.findElement(NextpageButtonProperty);
		NextpageButton.click();
		
		By webTableProperty1 = By.xpath(Properties.getProperty("EmployeeListEmployeeListWebTableProperty"));
		WebElement webTable1 = driver.findElement(webTableProperty1);

		By rowProperty1 = By.tagName("tr");
		List<WebElement> webTableRows1 = webTable1.findElements(rowProperty1);
		

		// Iterate through each row in the web table
		for (int rowIndex = 0; rowIndex < webTableRows1.size(); rowIndex++) {
			// Get the corresponding row in the Excel sheet
			Row excelRow = Page2.getRow(rowIndex);

			// going to a particular Row
			WebElement row1 = webTableRows1.get(rowIndex);

			// going to a particular Row and identify the number of its Corresponding Cells
			By rowofCellProperty = By.tagName("td");
			List<WebElement> rowOfCells = row1.findElements(rowofCellProperty);

			// Initialize cell index to start from column 14
			int cellIndex =0;

			// going to every Row of its Corresponding Cells
			for (int rowOfCellIndex = 0; rowOfCellIndex < rowOfCells.size(); rowOfCellIndex++) {
				// going to a Particular row of its Corresponding Cell
				WebElement rowOfCell = rowOfCells.get(rowOfCellIndex);

				if (excelRow == null) {
					excelRow = Page2.createRow(rowIndex);
				}

				// Get the corresponding cell in the Excel sheet
				Cell excelCell = excelRow.createCell(cellIndex++);

				// get the WebTable Data from the Row of Cell
				String webTableDataText = rowOfCell.getText();

				// Print the text of the current cell
				System.out.print(webTableDataText + " | ");

				// Write the text of the current cell to the Excel cell
				excelCell.setCellValue(webTableDataText);
			}
			System.out.println();

		}
		
		XSSFSheet Page3=workbook.createSheet("Sheet2");
		By NextpageButtonProperty1=By.xpath("/html/body/div[1]/div[3]/div[2]/div/form/div[5]/ul/li[7]/a");
		WebElement NextpageButton1=driver.findElement(NextpageButtonProperty1);
		NextpageButton1.click();
		
		By webTableProperty2 = By.xpath(Properties.getProperty("EmployeeListEmployeeListWebTableProperty"));
		WebElement webTable2 = driver.findElement(webTableProperty2);

		By rowProperty2 = By.tagName("tr");
		List<WebElement> webTableRows2 = webTable2.findElements(rowProperty2);
		

		// Iterate through each row in the web table
		for (int rowIndex =0; rowIndex < webTableRows2.size(); rowIndex++) {
			// Get the corresponding row in the Excel sheet
			Row excelRow = Page3.getRow(rowIndex);

			// going to a particular Row
			WebElement row1 = webTableRows2.get(rowIndex);

			// going to a particular Row and identify the number of its Corresponding Cells
			By rowofCellProperty = By.tagName("td");
			List<WebElement> rowOfCells = row1.findElements(rowofCellProperty);

			// Initialize cell index to start from column 14
			int cellIndex =0;

			// going to every Row of its Corresponding Cells
			for (int rowOfCellIndex = 0; rowOfCellIndex < rowOfCells.size(); rowOfCellIndex++) {
				// going to a Particular row of its Corresponding Cell
				WebElement rowOfCell = rowOfCells.get(rowOfCellIndex);

				if (excelRow == null) {
					excelRow = Page3.createRow(rowIndex);
				}

				// Get the corresponding cell in the Excel sheet
				Cell excelCell = excelRow.createCell(cellIndex++);

				// get the WebTable Data from the Row of Cell
				String webTableDataText = rowOfCell.getText();

				// Print the text of the current cell
				System.out.print(webTableDataText + " | ");

				// Write the text of the current cell to the Excel cell
				excelCell.setCellValue(webTableDataText);
			}
			System.out.println();

		}
		FileOutputStream FileOutput = new FileOutputStream(
				"./src/main/java/OrangeHrmApplicationTestResultFiles/EmployeeListResult.xlsx");
		workbook.write(FileOutput);
		
		
		
		

	}

}
