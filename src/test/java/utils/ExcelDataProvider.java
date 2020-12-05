package utils;

import java.io.IOException;
import java.util.concurrent.TimeUnit;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

public class ExcelDataProvider {

	WebDriver driver;
	private static Logger logger = LogManager.getLogger();
	
	@BeforeTest
	public void setUpLoginTest() {
		String projectPath = System.getProperty("user.dir");
		
		System.setProperty("webdriver.chrome.driver","C:\\Users\\koree\\Downloads\\chromedriver_win32\\chromedriver.exe");

		driver = new ChromeDriver();
	}

	@Test(dataProvider = "test1data")
	public void test1(String username, String password) throws Exception {
		System.out.println(username + "  |  " + password);

		driver = new ChromeDriver();
		logger.info("Open window");   
		
		driver.manage().deleteAllCookies();
		logger.info("Deleted cookies");
		
		driver.manage().window().maximize();
		logger.info("Maximixed the window");
		
		driver.manage().timeouts().implicitlyWait(15, TimeUnit.SECONDS);
		logger.info("Apply time");
		
		driver.get(" http://www.demo.guru99.com/V4/");
		logger.info("Launch browser");
		
	    driver.findElement(By.name("uid")).sendKeys(username); 
		logger.info("Enter username");
		
		driver.findElement(By.name("password")).sendKeys(password); 
		logger.info("Enter password");   
		
		Thread.sleep(3000);  
		driver.findElement(By.name("btnLogin")).click();
		logger.info("Click Login button");
		
		//driver.quit();
		logger.info("Close the browser");
		
	}
	
	@DataProvider(name = "test1data")
	public Object[][] getData() throws IOException {
		String excelPath = "C:\\Users\\koree\\Workspace\\Guru99BankProject\\excelguru\\excelguru.xls.xls.xlsx";
		Object data[][] = testData(excelPath, "Sheet1");
		return data;
	}

	public Object[][] testData(String excelPath, String SheetName) throws IOException {

		ExcelUtils excel = new ExcelUtils(excelPath, SheetName);

		int rowCount = excel.getRowCount();
		int colCount = excel.getColCount();

		Object data[][] = new Object[rowCount - 1][colCount];

		for (int i = 1; i < rowCount; i++) {
			for (int j = 0; j < colCount; j++) {

				String cellData = excel.getCellDataString(i, j);
				System.out.print(cellData + " | ");
				data[i - 1][j] = cellData;
			}

			// System.out.println();
		}
		return data;

	}
}