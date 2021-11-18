package AutomationPractice.PracticeCode;

import java.io.File;
import java.io.IOException;
import java.net.InetAddress;
import java.net.UnknownHostException;
import java.util.concurrent.TimeUnit;

import org.apache.commons.io.FileUtils;
import org.openqa.selenium.By;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import com.aventstack.extentreports.ExtentReports;
import com.aventstack.extentreports.ExtentTest;
import com.aventstack.extentreports.MediaEntityBuilder;
import com.aventstack.extentreports.Status;
import com.aventstack.extentreports.reporter.ExtentSparkReporter;
import com.aventstack.extentreports.reporter.configuration.Theme;

public class ExtentreportSelenium {
	WebDriver driver;
	ExtentReports extent;
    ExtentSparkReporter spark;
    ExtentTest test;
    
    
	@BeforeTest
	public void setup() throws UnknownHostException 
	{
     extent= new ExtentReports();
     spark = new ExtentSparkReporter("target/Selenium.html");
     spark.config().setTheme(Theme.DARK);
     extent.attachReporter(spark);
     test=extent.createTest("SeleniumTest");
     
     String myOsName=System.getProperty("os.name");
     extent.setSystemInfo("os", myOsName);
     String systemName=InetAddress.getLocalHost().getHostName();
     test.assignDevice(systemName);
     test.assignAuthor("Hemant");
     
     
       test.log(Status.INFO, "I am opening a browser");
	System.setProperty("webdriver.chrome.driver", "resource/chromedriver.exe");
	driver= new ChromeDriver();
	test.info("Started the browser");
	driver.get("http://automationpractice.com/index.php?controller=authentication&back=my-account");
	test.info("Started the automation practice site");
	driver.manage().window().maximize();
	driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
	test.pass("setup is complete");
   
     
    }
	
	
	@Test(priority = 1)
	 public void signup() throws IOException
     {
		try {
			test.info("I am logging in");
			   WebElement email=driver.findElement(By.id("email"));email.sendKeys("newaccount12345@gmail.com");
				WebElement pass=driver.findElement(By.id("passwd"));pass.sendKeys("newaccount12345");
				WebElement loginbtn=driver.findElement(By.id("SubmitLogin"));loginbtn.click();
				test.pass("I have logged in");
			
		} catch (Exception e) {
			test.fail("login had some issue"+ e.getMessage());
			test.fail(MediaEntityBuilder.createScreenCaptureFromPath(screenshot("screenshot1")).build());
		}
           
			
     }
	@Test(priority = 2)
	public void Select() throws IOException
	{
		try {
			test.info("I have started selection");
			WebElement women=driver.findElement(By.className("sf-with-ul"));women.click();
			test.pass("I have Completed Selection");
			test.pass("I have added by Saurab");
		} catch (Exception e) {
			test.fail("There is some problem in selection"+ e.getMessage());
			test.fail(MediaEntityBuilder.createScreenCaptureFromPath(screenshot("screenshot2")).build());
		}
		
	}
	
	@AfterTest
	public void teardown()
	{
	     
	
	try {
		Thread.sleep(2000);
	} catch (InterruptedException e) {
		// TODO Auto-generated catch block
		e.printStackTrace();
	}
	driver.quit();
	test.info("driver is closed");
	extent.flush();
	}
	
	public String screenshot(String screenShotName) throws IOException
	{
		TakesScreenshot ts = (TakesScreenshot)driver;
		File source = ts.getScreenshotAs(OutputType.FILE);
		String dest = System.getProperty("user.dir") +"\\Screens\\"+screenShotName+".png";
		File destination = new File(dest);
		FileUtils.copyFile(source, destination);
		return dest;        
	}

	
}


    
 
	
	


