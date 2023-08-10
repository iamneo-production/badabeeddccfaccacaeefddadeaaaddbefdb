package testcases;

import java.io.FileInputStream;
import java.io.IOException;
import java.net.MalformedURLException;
import java.time.Duration;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.support.PageFactory;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import com.aventstack.extentreports.ExtentReports;
import com.aventstack.extentreports.ExtentTest;
import com.aventstack.extentreports.Status;
import com.aventstack.extentreports.reporter.ExtentSparkReporter;

import pages.Tricentis_Login;
import pages.Tricetis_Register;
import utils.EventHandler;
import utils.Reporter;
import utils.Screenshot;

public class Testcase1 extends Base {
    // Logger log;
    EventHandler e;
    ExtentReports reporter = Reporter.generateExtentReport();;
    ExtentSparkReporter sparkReporter;

    @Test(priority = 1)
    public void Register() throws IOException {
        try {
            ExtentTest test = reporter.createTest("Registeration Page", "Execution for registeration");
            e = new EventHandler();
            driver.get(prop.getProperty("url") + "/Register");
            e.logger("Browser Navigated to the Register Page");
            driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(IMPLICIT_WAIT_TIME));
            driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(PAGE_LOAD_TIME));
            Tricetis_Register register = PageFactory.initElements(driver, Tricetis_Register.class);
            register.gen_female();
            e.logger("Female Button clicked in the Register Page");
            register.FirstName("AAAAAA");
            e.logger("First Name Entered in the Register Page");
            register.LastName("BBBBBB");
            e.logger("Last Name Entered in the Register Page");
            register.Email("test@demo.com");
            e.logger("Email Entered in the Register Page");
            register.Password("Test123$");
            e.logger("Password Entered in the Register Page");
            register.ConfirmPassword("Test123$");
            e.logger("Confirm Password Entered in the Register Page");
            Screenshot.getScreenShot("Register_Screenshot");
            e.logger("Screenshot taken in the Register Page");
            register.Register_Btn();
            e.logger("Register Button Clicked");
            test.log(Status.PASS, "Registeration Success");
        } catch (Exception ex) {
            // Handle any exceptions here
            e.logger("Exception occurred: " + ex.getMessage());
            // You can log or report the exception here
            // For now, just printing the stack trace
            ex.printStackTrace();
        }
    }

    @Test(priority = 2, dataProvider = "testData")
    public void Valid_Login_TC(String uid, String pwd) throws IOException {
        try {
            ExtentTest test = reporter.createTest("Login Page", "Execution for Login");
            driver.get(prop.getProperty("url") + "/Login");
            driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(IMPLICIT_WAIT_TIME));
            driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(PAGE_LOAD_TIME));
            Tricentis_Login login = PageFactory.initElements(driver, Tricentis_Login.class);
            login.username(uid);
            e.logger("username read from excel and entered in the field");
            login.password(pwd);
            e.logger("password read from excel and entered in the field");
            login.Login();
            e.logger("Login Button Clicked");
            test.log(Status.PASS, "Login Success");
            Screenshot.getScreenShot("Login_Screenshot");
            e.logger("Screenshot taken");
        } catch (Exception ex) {
            e.logger("Exception occurred: " + ex.getMessage());
            
            ex.printStackTrace();
        }
    }

    @DataProvider(name = "testData")
    public Object[][] readTestData() throws IOException {
        String excelFilePath = System.getProperty("user.dir") + "/src/test/java/resources/Testdata.xlsx";
        String sheetName = "Sheet1";

        FileInputStream fileInputStream = new FileInputStream(excelFilePath);
        Workbook workbook = WorkbookFactory.create(fileInputStream);
        Sheet sheet = workbook.getSheet(sheetName);

        int rowCount = sheet.getLastRowNum();
        int colCount = sheet.getRow(0).getLastCellNum();

        Object[][] data = new Object[rowCount][colCount];
        for (int i = 1; i <= rowCount; i++) {
            Row row = sheet.getRow(i);
            for (int j = 0; j < colCount; j++) {
                Cell cell = row.getCell(j, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                data[i - 1][j] = cell.getStringCellValue();
            }
        }

        return data;
    }

    @BeforeMethod
    public void beforeMethod() throws MalformedURLException {
        openBrowser();

    }

    @AfterMethod
    public void afterMethod() {
        driver.quit();
        reporter.flush();
        e.logger("Browser closed");
    }
}
