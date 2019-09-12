package com.cspire.Transfer;

import org.testng.annotations.Test;

import com.aventstack.extentreports.ExtentReports;
import com.aventstack.extentreports.ExtentTest;
import com.aventstack.extentreports.Status;
import com.aventstack.extentreports.reporter.ExtentHtmlReporter;

import org.testng.annotations.Test;
import org.testng.annotations.BeforeMethod;

import java.awt.AWTException;
import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.concurrent.TimeUnit;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Select;
import org.testng.annotations.AfterMethod;

public class Write implements path {
	WebDriver driver;
	ExtentReports extent;
	ExtentTest logger;
	
  @Test
  public void f() throws FileNotFoundException, IOException, AWTException {
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
	   driver.findElement(By.name("j0_73")).click();
	   driver.findElement(By.linkText("Count")).click();
	    driver.findElement(By.name("inventoryTree.children[4].selected")).click();
	    driver.findElement(By.name("inventoryTree.children[5].selected")).click();
	    driver.findElement(By.xpath("//td[4]/a/img")).click();
	    WebElement table1=driver.findElement(By.xpath("/html/body/table/tbody/tr[2]/td[3]/form/table[3]"));
      List<WebElement> rows=table1.findElements(By.tagName("tr"));
      int rowsize=rows.size();
      System.out.println(rowsize);
      XSSFWorkbook wbok= new XSSFWorkbook(new FileInputStream(excel_sheet));
      XSSFSheet wshet = wbok.getSheet("Mobile");
      
      FileOutputStream fout=new FileOutputStream(excel_sheet);
      
      outerloop:
      for(int i = 1;i<rows.size();i++)                                
      {
           List<WebElement> cols=rows.get(i).findElements(By.tagName("td"));
           
           if(cols.get(0).getText().equalsIgnoreCase(Device))  //enter the device name//
           {
          	 
          	 String quantity=cols.get(4).getText();
          	 System.out.println("AVAILABLE QUANTITY OF MOBILES AT STORE IS:"+ quantity);
          	 int quant=Integer.parseInt(quantity);
          	 List<WebElement> status=rows.get(i+1).findElements(By.tagName("td"));
          	 
          	 if((status.get(5).getText().equalsIgnoreCase("NEW")))
          			 
          	 {

          		 System.out.println(status.get(5).getText());
          		 System.out.println(status.get(6).getText());
          		if(status.get(6).getText().equalsIgnoreCase("READY"))
          		{
          			int k=0;
          	  for(int j=i+1;j<i+quant+1;j++)
          	  {
          		  List<WebElement> esn=rows.get(j).findElements(By.tagName("td"));
          		     
                    
                  	 //System.out.println(esn.get(2).getText());
                  	 String im=esn.get(2).getText();
                  	 String[] parts = im.split(":");
                  	 String part2 = parts[1];
                  	 wshet.createRow(k).createCell(0).setCellValue(part2);   
                  	 k=k+1;
                 }
          	  break outerloop;
          			}
          	  
          	    
           }
          	  	 
      }
          
	}
      
      wbok.write(fout);
      wbok.close();
      logger.log(Status.PASS, "ESN writed successfully to sheet");
      }

  
  @BeforeMethod
  public void beforeMethod()  {
	  ExtentHtmlReporter reporter=new ExtentHtmlReporter("./Reports/rep.html");
	  extent = new ExtentReports();
	  extent.attachReporter(reporter);
	  logger=extent.createTest("Write");
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

