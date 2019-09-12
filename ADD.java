
package com.cspire.Transfer;

import org.testng.annotations.Test;

import com.aventstack.extentreports.ExtentReports;
import com.aventstack.extentreports.ExtentTest;
import com.aventstack.extentreports.reporter.ExtentHtmlReporter;

import org.testng.annotations.BeforeMethod;

import java.awt.AWTException;
import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.concurrent.TimeUnit;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Select;
import org.testng.annotations.AfterMethod;


public class ADD implements path{
	
	ExtentReports extent;
	ExtentTest logger;
	WebDriver driver;
	
  @Test(priority=1)
  
  public void add() throws FileNotFoundException, IOException, AWTException {
	  driver.findElement(By.name("j_username")).clear();
	    driver.findElement(By.name("j_username")).sendKeys("");
	    driver.findElement(By.name("j_username")).clear();
	    driver.findElement(By.name("j_username")).sendKeys("ipsx");
	    driver.findElement(By.name("j_password")).clear();
	    driver.findElement(By.name("j_password")).sendKeys("ipsx");
	    driver.findElement(By.cssSelector("input[type=\"submit\"]")).click();
	    driver.findElement(By.cssSelector("a > font")).click();
	    new Select(driver.findElement(By.id("locs"))).selectByVisibleText(store);
	    driver.findElement(By.cssSelector("input[type=\"submit\"]")).click();
	    driver.findElement(By.name("j0_48")).click();
	    driver.findElement(By.name("j0_49")).click();
	   /* driver.findElement(By.linkText("Count")).click();
	    driver.findElement(By.name("inventoryTree.children[4].selected")).click();
	    driver.findElement(By.name("inventoryTree.children[5].selected")).click();
	    driver.findElement(By.xpath("//td[4]/a/img")).click();
	    WebElement table1=driver.findElement(By.xpath("/html/body/table/tbody/tr[2]/td[3]/form/table[3]"));*/
	    driver.findElement(By.linkText("New Transfer")).click();
	    new Select(driver.findElement(By.id("sNode2"))).selectByVisibleText("HATTIESBURG STORE");
	    driver.findElement(By.cssSelector("input[type=\"submit\"]")).click();
	    driver.findElement(By.xpath("//td[4]/a/img")).click();
	    driver.findElement(By.name("addLineSerCsn")).clear();
	    driver.findElement(By.name("addLineSerCsn")).sendKeys(Device);
	    driver.findElement(By.name("addLineEsnQuantity")).clear();
	    driver.findElement(By.name("addLineEsnQuantity")).sendKeys(Number);
	    driver.findElement(By.cssSelector("input[type=\"submit\"]")).click();
	    //XSSFWorkbook wbok= new XSSFWorkbook(new FileInputStream("D:\\demo.xlsx"));
	    //XSSFSheet wshet = wbok.getSheet("shubham");
	    XSSFWorkbook wbok= new XSSFWorkbook(new FileInputStream(excel_sheet));
	    XSSFSheet wshet = wbok.getSheet("Mobile");
	    XSSFRow row=wshet.getRow(0);
	    Robot robot = new Robot();
	   
	    
	    for(int i=0;i<Integer.parseInt(Number);i++)
	    {
	    	row = wshet.getRow(i);
	    	for(int j=0;j<=2;j++)
	    		 
	    	{
	    	 
	    	  try{
	    	 
	    		driver.findElement(By.id("newEsnId")).clear();
	  		    driver.findElement(By.id("newEsnId")).sendKeys(row.getCell(0).getStringCellValue());
	    	 
	    	    break;
	    	 
	    	}
	    	
	    	 
	    	catch(Exception e)
	    	{
	    	 
	    	System.out.println(e.getMessage());
	    	 
	    	}
	    	 
	    	}
	   
	    
	    robot.keyPress(KeyEvent.VK_ENTER); //press enter key
	    robot.keyRelease(KeyEvent.VK_ENTER);
	    
	    }
	    driver.findElement(By.xpath("//td[2]/a/img")).click();
	    driver.findElement(By.xpath("//td[2]/a/img")).click();
	    
	    driver.findElement(By.linkText("Logout")).click();   
	    //
	    recieve();
	  
  }
  
  public void recieve() throws FileNotFoundException, IOException, AWTException {
	  driver.findElement(By.name("j_username")).clear();
	    driver.findElement(By.name("j_username")).sendKeys("");
	    driver.findElement(By.name("j_username")).clear();
	    driver.findElement(By.name("j_username")).sendKeys("ipsx");
	    driver.findElement(By.name("j_password")).clear();
	    driver.findElement(By.name("j_password")).sendKeys("ipsx");
	    driver.findElement(By.cssSelector("input[type=\"submit\"]")).click();
		//driver.findElement(By.cssSelector("a > font")).click();
		 //driver.findElement(By.name("j0_49")).click();
		 //driver.findElement(By.name("j0_50")).click();
		 driver.findElement(By.linkText("Receivable")).click();
		   driver.findElement(By.cssSelector("u")).click();
		   driver.findElement(By.cssSelector("img[alt=\"Receive items\"]")).click();
		   driver.findElement(By.name("numberToReceive")).clear();
		   driver.findElement(By.name("numberToReceive")).sendKeys(Number);
		   driver.findElement(By.xpath("//div[@id='button_row']/table/tbody/tr/td[2]/a/img")).click();
		   XSSFWorkbook wbok= new XSSFWorkbook(new FileInputStream(excel_sheet));
		    XSSFSheet wshet = wbok.getSheet("Mobile");
		    XSSFRow row=wshet.getRow(0);
		    Robot robot = new Robot();
		    int i=0;
		    for(i=0;i<Integer.parseInt(Number);i++)
		    {
		    	row = wshet.getRow(i);
		    	for(int j=0;j<=2;j++)
		    		 
		    	{
		    	 
		    	  try{
		    		  driver.findElement(By.id("newEsnId"+"["+String.valueOf(i)+"]")).clear();
				        driver.findElement(By.id("newEsnId"+"["+String.valueOf(i)+"]")).sendKeys(row.getCell(0).getStringCellValue());
		    		
		    	 
		    	    break;
		    	 
		    	}
		    	
		    	 
		    	catch(Exception e)
		    	{
		    	 
		    	System.out.println(e.getMessage());
		    	 
		    	}
		    	 
		    	}
		    	
		        robot.keyPress(KeyEvent.VK_ENTER); //press enter key
			    robot.keyRelease(KeyEvent.VK_ENTER);
		    }
  }
  
  
  @BeforeMethod
  public void beforeMethod() {
	  ExtentHtmlReporter reporter=new ExtentHtmlReporter("./Reports/rep.html");
	  extent = new ExtentReports();
	  reporter.setAppendExisting(true);
	  extent.attachReporter(reporter);
	  logger=extent.createTest("ADD");
	 
	  System.setProperty("webdriver.chrome.driver",driver_path);
      String url = "https://testappa.cspire.net/ips/";
      driver = new ChromeDriver();
      driver.get(url);
     
      driver.manage().window().maximize();
      driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
      
  }

  @AfterMethod
  public void afterMethod() {
	  
	  driver.close();
		driver.quit();
		extent.flush();
  }

}
