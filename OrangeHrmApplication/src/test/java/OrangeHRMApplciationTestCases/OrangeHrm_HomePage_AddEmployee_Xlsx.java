package OrangeHRMApplciationTestCases;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
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



public class OrangeHrm_HomePage_AddEmployee_Xlsx extends BaseTest {
	
	FileInputStream OrangeHrmApplicationpropertiesFile;
	Properties Properties;
	
	@Test(priority=1,description=" Validating OrangeHRM Application Add Employee")
	public void addEmployeeValiation() throws IOException, InterruptedException {
		FileInputStream addEmployeeTestData = new FileInputStream(
				"./src/main/java/OrangeHrmApplicationTestDataFiles/AddEmployeeTestData.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(addEmployeeTestData);
		XSSFSheet addEmployeeTestDatasheet = workbook.getSheet("Add_Employee_Validtion");
		
		OrangeHrmApplicationpropertiesFile = new FileInputStream("./src/main/java/Config/OrangeHrmApplication.Properties");
		Properties= new Properties();
		Properties.load(OrangeHrmApplicationpropertiesFile);
		

		Row row = addEmployeeTestDatasheet.getRow(1);

		Cell expectedLoginPanelRowofCell = row.getCell(0);

		String ExpctedLoginPanelText = expectedLoginPanelRowofCell.getStringCellValue();
		Log.info("This Expected Login Panel Text Is:-" + ExpctedLoginPanelText);

		By loginPanelproperty = By.id(Properties.getProperty("LogInPageLoginPanelProperty"));
		WebElement loginPanel = driver.findElement(loginPanelproperty);
		String actualLoginPanelText = loginPanel.getText();
		Log.info("The Actual Login Panel Text Is:- " + actualLoginPanelText);
		Cell Actual_LogInPageTextRowOfCell = row.createCell(2);
		Actual_LogInPageTextRowOfCell.setCellValue(actualLoginPanelText);
		Cell loginTextTestResult = row.createCell(3);

		if (actualLoginPanelText.equals(ExpctedLoginPanelText)) {
			Log.info("Successfully Navigated to OrangeHrm Login Page:-Pass");
			loginTextTestResult.setCellValue("Pass");
		} else {
			Log.info("Failed to Navigate to OrangeHrm Login Page:-Fail");
			loginTextTestResult.setCellValue("Fail");
		}
		System.out.println();

		Cell LogInPageTitle = row.getCell(4);
		String expected_LogInPageTitle = LogInPageTitle.getStringCellValue();
		Log.info("The Expected Title of the OrangeHrm Login Page is:-" + expected_LogInPageTitle);

		String actual_LoginPageTitle = driver.getTitle();
		Log.info("The Actual Title of the OrangeHrm Login Page is:-" + actual_LoginPageTitle);
		Cell actual_LoginPageTitleRowofCell = row.createCell(5);
		actual_LoginPageTitleRowofCell.setCellValue(actual_LoginPageTitle);
		Cell titleTestResult = row.createCell(6);

		if (actual_LoginPageTitle.equalsIgnoreCase(expected_LogInPageTitle)) {
			Log.info("The Title of the OrangeHrm Page is Matched:-Pass");
			titleTestResult.setCellValue("Pass");
		} else {
			Log.info("The Title of the OrangeHrm Page is Not Matched:-Fail");
			titleTestResult.setCellValue("Fail");
		}
		System.out.println();

		Cell validUsernameRowOfCell = row.getCell(7);
		String usernameTestData = validUsernameRowOfCell.getStringCellValue();

		By usernameProperty = By.id(Properties.getProperty("LogInPageUserNameProperty"));
		WebElement username = driver.findElement(usernameProperty);
		username.sendKeys(usernameTestData);

		Cell validPasswordRowOfCell = row.getCell(8);
		String paswordTestData = validPasswordRowOfCell.getStringCellValue();

		By passwordProperty = By.id(Properties.getProperty("LogInPagePasswordProperty"));
		WebElement password = driver.findElement(passwordProperty);
		password.sendKeys(paswordTestData);

		By loginbuttonProperty = By.id(Properties.getProperty("LogInPageLoginButtonProperty"));
		WebElement loginbutton = driver.findElement(loginbuttonProperty);
		loginbutton.click();

		Cell expectedHomePageTextRowofCell = row.getCell(9);
		String expectedHomePageText = expectedHomePageTextRowofCell.getStringCellValue();
		Log.info("The Expected OrangeHrm Home Page Text is:-" + expectedHomePageText);

		By welcomeAdminProperty = By.id(Properties.getProperty("HomePageWelcomAdminProperty"));
		WebElement welcomeadmin = driver.findElement(welcomeAdminProperty);
		String actualHomePageText = welcomeadmin.getText();
		Cell actualHomePageTextRowOfCell = row.createCell(10);
		actualHomePageTextRowOfCell.setCellValue(actualHomePageText);
		Log.info("The Actual OrangeHrm Home Page Text is:-" + actualHomePageText);

		Cell homePageTestResult = row.createCell(11);
		if (actualHomePageText.contains(expectedHomePageText)) {
			Log.info("Successfully Navigated to OrangeHrm HomePage:-Pass");
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

		By addEmployeeProperty = By.id(Properties.getProperty("HomePageAddEmployeeProperty"));
		WebElement addEmployee = driver.findElement(addEmployeeProperty);
		addEmployee.click();

		Cell expected_AddEmployeePageTextRowOfCell = row.getCell(12);
		String expected_AddEmployeePageText = expected_AddEmployeePageTextRowOfCell.getStringCellValue();
		Log.info("The Expected Add Employee Page Text is :-" + expected_AddEmployeePageText);

		By AddEmployeePageTextProperty = By.xpath(Properties.getProperty("AddEmployeePageAddEmplpyeeText"));
		WebElement AddEmployeePageText = driver.findElement(AddEmployeePageTextProperty);
		String actual_AddEmployeePageText = AddEmployeePageText.getText();
		Log.info("The Actual Add Employee Page Text is:-" + actual_AddEmployeePageText);
		Cell actual_AddEmployeePageTextRowOfCell = row.createCell(13);
		actual_AddEmployeePageTextRowOfCell.setCellValue(actual_AddEmployeePageText);

		Cell addEmployee_TestResult = row.createCell(14);
		if (actual_AddEmployeePageText.equals(expected_AddEmployeePageText)) {
			Log.info("Successfully Navigated to Add Employee Page:-Pass");
			addEmployee_TestResult.setCellValue("Pass");
		} else {
			Log.info("Failed to Navigate to Add Employee Page:-Fail");
			addEmployee_TestResult.setCellValue("Fail");
		}
		System.out.println();

		int rowCount = addEmployeeTestDatasheet.getLastRowNum();

		for (int rowIndex = 1; rowIndex <= rowCount; rowIndex++) {
			Row row1 = addEmployeeTestDatasheet.getRow(rowIndex);
			Cell cell = row1.getCell(15);
			if (cell != null && !cell.toString().isEmpty()) {

				Cell Expected_FirstNameTextDataCell = row1.getCell(15);
				String Expected_FirstNameTextData = Expected_FirstNameTextDataCell.getStringCellValue();

				Cell Expected_MiddleNameTextDataCell = row1.getCell(16);
				String Expected_MiddleNameTextData = Expected_MiddleNameTextDataCell.getStringCellValue();

				Cell Expected_LastNameTextDataCell = row1.getCell(17);
				String Expected_LastNameTextData = Expected_LastNameTextDataCell.getStringCellValue();

				By EmployeeIdProperty = By.id(Properties.getProperty("AddEmployeePageEmployeeIdProperty"));
				WebElement EmployeeId = driver.findElement(EmployeeIdProperty);
				String expected_EmployeeId = EmployeeId.getAttribute("Value");
				Cell expected_EmployeeIdRowOFCell = row1.createCell(18);
				expected_EmployeeIdRowOFCell.setCellValue(expected_EmployeeId);

				By firstNameProperty = By.id(Properties.getProperty("AddEmployeePageFirstNameProperty"));
				WebElement firstName = driver.findElement(firstNameProperty);
				firstName.sendKeys(Expected_FirstNameTextData);

				By middleNameProperty = By.id(Properties.getProperty("AddEmployeePageMiddleNameProperty"));
				WebElement middleName = driver.findElement(middleNameProperty);
				middleName.sendKeys(Expected_MiddleNameTextData);

				By lastNameProperty = By.id(Properties.getProperty("AddEmployeePageLastNameProperty"));
				WebElement lastName = driver.findElement(lastNameProperty);
				lastName.sendKeys(Expected_LastNameTextData);

				By saveButtonProperty = By.id(Properties.getProperty("AddEmployeePageSaveButtonProperty"));
				WebElement saveButton = driver.findElement(saveButtonProperty);
				saveButton.click();

				Thread.sleep(2000);
				Cell expected_PersonalDetailsPageTextCell = row1.getCell(19);
				String expected_PersonalDetailsPageText = expected_PersonalDetailsPageTextCell.getStringCellValue();

				By PersonalDetailsPageTextProperty = By.xpath(Properties.getProperty("personalDetialsPageText"));
				WebElement PersonalDetailsPageText = driver
						.findElement(PersonalDetailsPageTextProperty);
				String actual_PersonalDetailsPageText = PersonalDetailsPageText.getText();

				Cell actual_PersonalDetailsPageTextRowofCell = row1.createCell(20);
				actual_PersonalDetailsPageTextRowofCell.setCellValue(actual_PersonalDetailsPageText);

				Cell PersonalDetailsPageTextResult = row1.createCell(21);

				if (actual_PersonalDetailsPageText.equals(expected_PersonalDetailsPageText)) {
					Log.info("Successfully Navigated to OrangeHrm PersonalDetailsPage:-Pass");
					PersonalDetailsPageTextResult.setCellValue("Pass");
				} else {
					Log.info("Failed to  Navigate to OrangeHrm PersonalDetailsPage:-Fail");
					PersonalDetailsPageTextResult.setCellValue("Fail");
				}

				By firstNamePersonalPageProperty = By.id(Properties.getProperty("PersonalDetailsPageFirstNameProperty"));
				WebElement firstNamePersonalPage = driver.findElement(firstNamePersonalPageProperty);
				String actual_firstNamePersonalPage = firstNamePersonalPage.getAttribute("Value");
				Cell actual_firstNamePersonalPagerRowOfCell = row1.createCell(22);
				actual_firstNamePersonalPagerRowOfCell.setCellValue(actual_firstNamePersonalPage);
				Cell firstNameTestResult = row1.createCell(23);

				if (actual_firstNamePersonalPage.equals(Expected_FirstNameTextData)) {
					Log.info("The FirstName is matched:-Pass");
					firstNameTestResult.setCellValue("Pass");
				} else {
					Log.info("The First Name  is Not matched:-Fail");
					firstNameTestResult.setCellValue("Fail");
				}

				By middleNamePersonalPageProperty = By.id(Properties.getProperty("PersonalDetailsPageMiddleNameProperty"));
				WebElement middleNamePersonalPage = driver.findElement(middleNamePersonalPageProperty);
				String actual_middleNamePersonalPage = middleNamePersonalPage.getAttribute("Value");
				Cell actual_middleNamePersonalPageRowOfCell = row1.createCell(24);
				actual_middleNamePersonalPageRowOfCell.setCellValue(actual_middleNamePersonalPage);

				Cell middleNameTestResult = row1.createCell(25);

				if (actual_middleNamePersonalPage.equals(Expected_MiddleNameTextData)) {
					Log.info("The Middle Name is matched:-Pass");
					middleNameTestResult.setCellValue("Pass");
				} else {
					Log.info("The Middle Name  is Not matched:-Fail");
					middleNameTestResult.setCellValue("Fail");
				}

				By lastNamePersonalPageProperty = By.id(Properties.getProperty("PersonalDetailsPageLastNameProperty"));
				WebElement lastNamePersonalPage = driver.findElement(lastNamePersonalPageProperty);
				String actual_lastNamePersonalPage = lastNamePersonalPage.getAttribute("Value");
				Cell actual_lastNamePersonalPageRowOfCell = row1.createCell(26);
				actual_lastNamePersonalPageRowOfCell.setCellValue(actual_lastNamePersonalPage);

				Cell lastNameTestResult = row1.createCell(27);

				if (actual_lastNamePersonalPage.equals(Expected_LastNameTextData)) {
					Log.info("The last Name is matched:-Pass");
					lastNameTestResult.setCellValue("Pass");
				} else {
					Log.info("The Last Name  is Not matched:-Fail");
					lastNameTestResult.setCellValue("Fail");
				}

				By EmployeeIdPersonalPageProperty = By.id(Properties.getProperty("PersonalDetailsPageEmployeeIdProperty"));
				WebElement EmployeeIdPersonalPage = driver.findElement(EmployeeIdPersonalPageProperty);
				String actual_EmployeeIdPersonalPage = EmployeeIdPersonalPage.getAttribute("Value");
				Cell actual_EmployeeIdPersonalPageRowOfCell = row1.createCell(28);
				actual_EmployeeIdPersonalPageRowOfCell.setCellValue(actual_EmployeeIdPersonalPage);

				Cell EmployeeIdTestResult = row1.createCell(29);

				if (actual_EmployeeIdPersonalPage.equals(expected_EmployeeId)) {
					Log.info("The Employee Id is matched:-Pass");
					EmployeeIdTestResult.setCellValue("Pass");
				} else {
					Log.info("The Employee Id   is Not matched:-Fail");
					EmployeeIdTestResult.setCellValue("Fail");
				}
				System.out.println();

				WebElement welcomeadmin1 = driver.findElement(welcomeAdminProperty);
				welcomeadmin1.click();
				Thread.sleep(2000);

				By logOutProperty = By.linkText(Properties.getProperty("HomePageLogoutProperty"));
				WebElement logOut = driver.findElement(logOutProperty);
				logOut.click();

				Cell expectedLoginPanel1RowofCell = row.getCell(30);
				String expectedLoginPanel1 = expectedLoginPanel1RowofCell.getStringCellValue();

				WebElement loginPanel1 = driver.findElement(loginPanelproperty);
				String actualLoginPanelText1 = loginPanel1.getText();
				Cell actualLoginPanelText1RowOfCell = row.createCell(31);
				actualLoginPanelText1RowOfCell.setCellValue(actualLoginPanelText1);

				Cell loginTextTestResult1 = row.createCell(32);
				if (actualLoginPanelText1.equals(expectedLoginPanel1)) {
					Log.info("Successfully Navigated to OrangeHrm Login Page:-Pass");
					loginTextTestResult1.setCellValue("Pass");
				} else {
					Log.info("Failed to Navigate to OrangeHrm Login Page:-Fail");
					loginTextTestResult1.setCellValue("Fail");
				}
				WebElement username1 = driver.findElement(usernameProperty);
				username1.sendKeys(usernameTestData);

				WebElement password1 = driver.findElement(passwordProperty);
				password1.sendKeys(paswordTestData);

				WebElement loginbutton1 = driver.findElement(loginbuttonProperty);
				loginbutton1.click();

				Thread.sleep(2000);

				WebElement pim1 = driver.findElement(pimProperty);
				mouseHoverOperation.moveToElement(pim1).build().perform();

				WebElement addEmployee1 = driver.findElement(addEmployeeProperty);
				addEmployee1.click();

			}


		}

		FileOutputStream testResultFile = new FileOutputStream(
				"./src/main/java/OrangeHrmApplicationTestResultFiles/AddEmployeeTestResult.xlsx");
		workbook.write(testResultFile);

	}
	
		
	}
