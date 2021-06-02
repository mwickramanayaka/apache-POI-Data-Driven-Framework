package util;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import io.github.bonigarcia.wdm.WebDriverManager;

public class ExcelDataProvider {

	WebDriver driver = null;

	@BeforeTest
	public void setUpTest() {
		//initializing and starting the browser
		WebDriverManager.chromedriver().setup();
		driver = new ChromeDriver();
		driver.manage().window().maximize();

	}

	@Test(dataProvider="test1data")
	public void test1(String USERNAME, String PASSWORD) throws InterruptedException {
		System.out.println(USERNAME + "|" + PASSWORD);
		
		driver.get("https://vle.bit.lk/login/index.php");
		driver.findElement(By.id("username")).sendKeys(USERNAME);
		driver.findElement(By.id("password")).sendKeys(PASSWORD);
		Thread.sleep(2000);
		
	}
	
	@AfterTest
	public void tearDown() {
		driver.quit();
		
	}

	@DataProvider(name = "test1data")
	public static Object[][] getData() {

		String projectPath = System.getProperty("user.dir");
		String excelPath = projectPath + "/Excel Files/data1.xlsx";

		Object data[][] = testData(excelPath, "Sheet1");
		return data;

	}
 
	public static Object[][] testData(String exelPath, String sheetName) {

		ExcelUtilTest excel = new ExcelUtilTest (exelPath, sheetName);

		int rowCount = excel.getRowCount();
		int colCount = excel.getColCount();  

		Object data[][] = new Object[rowCount-1][colCount];

		for (int i = 1; i < rowCount; i++) {
			for(int j = 0; j < colCount; j++) {
				String cellData = excel.GetCellDataString(i, j);
				//System.out.print(cellData + " | ");
				data[i-1][j] = cellData;

			}
			//System.out.println();
		}
		return data;
	}  

}
