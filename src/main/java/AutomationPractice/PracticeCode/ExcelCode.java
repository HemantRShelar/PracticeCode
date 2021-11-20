package AutomationPractice.PracticeCode;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.concurrent.TimeUnit;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

public class ExcelCode {

	public static void main(String[] args) throws IOException {
		
		File file=new File("resource/Book1.xlsx");
		FileInputStream fis= new FileInputStream(file);
		
		   XSSFWorkbook excel = new XSSFWorkbook(fis);
			XSSFSheet sheet= excel.getSheet("Sheet1");
			
			String x= sheet.getRow(0).getCell(0).getStringCellValue();
			String y= sheet.getRow(0).getCell(1).getStringCellValue();
			System.out.println(x);
			System.out.println(y);
			
			String a=sheet.getRow(1).getCell(0).getStringCellValue();
			String b=sheet.getRow(1).getCell(1).getStringCellValue();
			System.out.println(a);
			System.out.println(b);
			System.setProperty("webdriver.chrome.driver","resource/chromedriver.exe");
			 WebDriver driver = new ChromeDriver();
			driver.get("https://opensource-demo.orangehrmlive.com/");
			driver.manage().window().maximize();
			driver.manage().timeouts().implicitlyWait(20, TimeUnit.SECONDS);
			
			driver.findElement(By.id("txtUsername")).sendKeys(a);
			driver.findElement(By.id("txtPassword")).sendKeys(b);
			driver.findElement(By.name("Submit")).click();
			
			sheet.getRow(1).createCell(2).setCellValue("PASS");
			FileOutputStream fos= new FileOutputStream(file);
			excel.write(fos);
			
			
			
     fis.close();
     fos.close();
     excel.close();
     driver.quit();
		}
       
		
		

	}


