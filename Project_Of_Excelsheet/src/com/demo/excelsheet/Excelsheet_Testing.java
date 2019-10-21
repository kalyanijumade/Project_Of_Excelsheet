package com.demo.excelsheet;

import org.testng.annotations.Test;
import org.testng.AssertJUnit;
import org.testng.annotations.Test;
import org.testng.asserts.SoftAssert;

import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

import org.testng.annotations.BeforeSuite;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.List;
import java.util.logging.Logger;

import org.apache.commons.io.FileUtils;
import org.apache.commons.io.FilenameUtils;
import org.openqa.selenium.By;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.testng.ITestResult;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.AfterSuite;
import org.testng.annotations.BeforeMethod;

public class Excelsheet_Testing {
	static ExtentTest test;
	
    static WebDriver driver;
    
    static ExtentReports report;
      
    SoftAssert soft = new SoftAssert();
    
    @BeforeMethod
    public void ReportGenerat() {
    	
    	report = new ExtentReports(System.getProperty("user.dir")+"/test-output/TestReportResult.html",true);
    	report.addSystemInfo("Host Name", "kalyani")
               .addSystemInfo("Environment", "Automation Testing")
               .addSystemInfo("User Name", "Seleneium Test");
    	report.loadConfig(new File(System.getProperty("user.dir")+"\\extent-config.xml"));
    }
	
	 @Test(priority =1)
	  public void Campare_Actual_with_Expected()throws Throwable
	 {
		 test = report.startTest("Name of Test Case", "One");
		FileInputStream fis = new FileInputStream("D:\\Selenium Software\\excelsheet\\expectedsheet.xls");  
		
		Workbook workbook1 = Workbook.getWorkbook(fis);
		
		driver = new FirefoxDriver();
		
		driver.get("file:///D:/Selenium%20Software/Selenium%20Software/Offline%20Website/pages/examples/users.html");
		
		driver.manage().window().maximize();
		
		Sheet sheet = workbook1.getSheet(0);
		
		int rows =sheet.getRows();
		int columns =sheet.getColumns();
		
	    for(int i=2;i<columns;i++){
			
			for(int j=2;j<rows+1;j++) {
				
				System.out.println("***************************************************************************************");
				
				Cell cell = sheet.getCell((i-1),(j-1));
				String Exceltable = cell.getContents();
				
				WebElement list = driver.findElement(By.xpath("/html/body/div[1]/div[1]/section[2]/div/div/div/div[2]/table/tbody/tr[" + (j) + "]/td[" + (i) + "]"));
				String Webtable =list.getText();
				
				System.out.println("Excelsheet Table :"+Exceltable);
				System.out.print("Web Table : "+Webtable);
				
				if(Exceltable.equals(Webtable)) {
					
					System.out.println("  Data Matched......");
					test.log(LogStatus.PASS,"User Data ["+ Exceltable +"] and Web Table [" + Webtable +"] is matching " );
					AssertJUnit.assertEquals(Webtable,Exceltable);
				}
				else {
					
					System.out.println("  Data Unmatched......");
					test.log(LogStatus.FAIL,"User Data ["+ Exceltable +"] and Web Table [" + Webtable +"] is unmatching " );
					AssertJUnit.assertEquals(Webtable,Exceltable);
				}
				
			}
	    }
		
	soft.assertAll();
	 }	
	

	 

@Test(priority=2)
     public static void ScreenShort_Method() throws Exception {
	
     File screenshot =((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
     FileUtils.copyFile(screenshot, new File("D:\\Selenium Software\\ScreenShortFile\\TestScreenShort.png"));
  }
 



@AfterMethod
public void getResult(ITestResult result) {
	
	if(result.getStatus() == ITestResult.FAILURE) {
		
		test.log(LogStatus.FAIL, " Test Case Failed is" +result.getName());
		test.log(LogStatus.FAIL, " Test Case Failed is" +result.getThrowable());
	}
	else if(result.getStatus() == ITestResult.SKIP) {
		
		test.log(LogStatus.SKIP, " Test Case Skipped is" +result.getName());
	}
	  report.flush();
}
 

}
