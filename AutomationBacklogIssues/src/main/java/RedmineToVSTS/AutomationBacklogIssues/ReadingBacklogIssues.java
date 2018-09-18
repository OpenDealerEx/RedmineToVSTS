package RedmineToVSTS.AutomationBacklogIssues;

import java.io.IOException;

import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;

import UtilFiles.ReadAndWriteExcelData;

public class ReadingBacklogIssues 


{
	
	public static void main(String[] args) throws InterruptedException,  IOException {
		
		
		System.setProperty("webdriver.chrome.driver", "C:\\Browser\\chromedriver.exe");
		WebDriver driver = new ChromeDriver();
			
	    ReadAndWriteExcelData excel= new ReadAndWriteExcelData("C:\\Users\\mgopalan\\Desktop\\VSTSREDMINE.xlsx");
		  				
		driver.get("https://opendealerexchange.visualstudio.com");
		driver.findElement(By.name("loginfmt")).sendKeys("mgopalan@opendealerexchange.com");
		
		driver.findElement(By.xpath("//input[@type='submit']")).click();
	    driver.findElement(By.name("passwd")).sendKeys("Waterloo1979*");
	    Thread.sleep(2000);
	    driver.findElement(By.xpath("//input[@type='submit']")).click();
	    
	    Thread.sleep(20000);
	    driver.findElement(By.xpath("//input[@type='button']")).click();
	   
		Thread.sleep(4000);
	   
	   	driver.navigate().to("https://ode-qa.visualstudio.com/Test%20Project/_backlogs/backlog/Test%20Project%20Team/Stories");
	    
	    Thread.sleep(3000);
	    		    		   
	    for( int i=1; i< 4;i++) {
	    driver.findElement(By.name("New Work Item")).click();
	 
	    String ticketNum  =excel.getData(0,i,0);
	    String project    =excel.getData(0,i,1);
	    String tracker    =excel.getData(0,i,2);
	    String status     =excel.getData(0,i,3);
	 	String subject    =excel.getData(0,i,4);
		String assignee   =excel.getData(0,i,5);
		String date       =excel.getData(0,i,6);
		String description=excel.getData(0,i,7);
		
		Thread.sleep(2000);
	  
		driver.findElement(By.className("ms-TextField-field")).sendKeys(subject);
	    driver.findElement(By.className("ms-Button-label")).click();
	   
	    Thread.sleep(3000);
		driver.findElement(By.xpath("//*[@id=\"row_vss_2_0\"]/div[4]/div[2]/a[contains(@class,'work-item-title-link') and contains(text(),'"+subject+"')]")).click();
	    
	    //driver.findElement(By.xpath("//input[contains(text(),'"+subject+"')]")).click();
		Thread.sleep(2000);
		
					
		if (status.equalsIgnoreCase("Closed")) {
				
			   // Thread.sleep(2000);				
			    WebElement state = driver.findElement(By.xpath("//div[@class='combo input-text-box list drop']"));
							
				Actions action = new Actions(driver);
				       
			    action.moveToElement(state).click();
		       
		      	action.keyDown(Keys.CONTROL).sendKeys(Keys.ARROW_UP).doubleClick();	
		       
		        action.perform();
		      
		  					
		}else  {
			
			   
				WebElement state = driver.findElement(By.xpath("//div[@class='combo input-text-box list drop']"));
			     
		        Actions action = new Actions(driver);
		        //action.click();
		        action.moveToElement(state).click();
		  	       
		        action.keyDown(Keys.CONTROL).sendKeys(Keys.ARROW_UP).sendKeys(Keys.ARROW_UP).doubleClick();
		    
		        action.perform();
		      
		       Thread.sleep(1000);
		      
	      
		}
	
		Thread.sleep(1000);
		driver.findElement(By.className("tag-box-selectable")).click();
		Thread.sleep(2000);
	    driver.findElement(By.className("tags-input")).sendKeys(ticketNum);
	    		 			
	    driver.findElement(By.className("tags-input")).sendKeys(Keys.RETURN);
	    Thread.sleep(1000);
	    driver.findElement(By.className("tags-input")).sendKeys(ticketNum);
	    Thread.sleep(1000);
		driver.findElement(By.className("tags-input")).sendKeys(Keys.RETURN);
		Thread.sleep(1000);
		driver.findElement(By.className("tags-input")).sendKeys(tracker);  
		Thread.sleep(1000);
		driver.findElement(By.className("tags-input")).sendKeys(Keys.RETURN);
		driver.findElement(By.className("tags-input")).sendKeys(project);
		Thread.sleep(1000);
		driver.findElement(By.className("tags-input")).sendKeys(Keys.RETURN);
		driver.findElement(By.className("tags-input")).sendKeys(assignee);
		Thread.sleep(1000);
		
		driver.findElement(By.className("tags-input")).sendKeys(Keys.RETURN);
			
		driver.findElement(By.className("tags-input")).sendKeys(date);
		
		Thread.sleep(1000);
		
		
		driver.findElement(By.xpath("//div[@class='richeditor-editarea']")).click();
		driver.switchTo().frame(0).findElement(By.xpath("html/body")).sendKeys(description);
		Thread.sleep(2000);  
		driver.switchTo().defaultContent(); 
		
		Thread.sleep(2000); 
			WebElement element = driver.findElement(By.xpath("//*[starts-with(@id,'vss_')]/div/div/div/div[1]/div[1]/div[2]/span"));
		      //WebElement element = driver.findElement(By.xpath("//*[@id=\"vss_4\"]/div/div/div/div[1]/div[1]/div[2]/span"));  	
			String	innertText=element.getText();
			System.out.println(innertText); 
			excel.writeData(0,i,8,innertText);

		driver.findElement(By.xpath("//span[contains(text(),'Save & Close')]")).click();
		Thread.sleep(2000);
	  
	 
	}
	   
	    driver.close();
	    driver.quit();   
	   
	}

}
