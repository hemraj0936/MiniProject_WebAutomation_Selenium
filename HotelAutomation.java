package com.selenium.project;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Properties;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.StaleElementReferenceException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.firefox.FirefoxOptions;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.Test;

public class HotelAutomation {
	WebDriver driver = null;
	WebDriverWait wait = null;
	List<WebElement> Hotel = null;
	List<WebElement> price = null;
	List<WebElement> locate = null;
	//public Properties prop = null;
	
	static Properties prop = readProperties();
	public static Properties readProperties() {
		File f = new File("config.properties");
		  
		FileInputStream fileInput = null;
		
		try {
			fileInput = new FileInputStream(f);
		} 
		catch (FileNotFoundException e) {
			e.printStackTrace();
		}
		
		Properties prop = new Properties();
		
		try {
			prop.load(fileInput);
		} 
		catch (IOException e) {
			e.printStackTrace();
		}
		
		return prop;
	}
	
	@BeforeMethod
    //use to start the browser
	public void OpenBrowser() {
		
		//taking Key(1-2) from the properties file
		String browser = prop.getProperty("browser");
		
		if(browser.equalsIgnoreCase("firefox")) 
		{
			//taking Key(4) from the properties file
			System.setProperty("webdriver.gecko.driver", "C:\\Users\\hemra\\eclipse-workspace\\Miniproject\\driver\\geckodriver.exe");
			FirefoxOptions fo = new FirefoxOptions();
			
			fo.addArguments("--disable-infobars");
            fo.addPreference("dom.webnotifications.enabled", false);
		    fo.addPreference("geo.enabled", false);
		    driver = new FirefoxDriver(fo);
		    driver.manage().window().maximize();
			}
		else
        {
	       //taking Key(5) from the properties file
	       System.setProperty("webdriver.chrome.driver",prop.getProperty("chromepath" ));
	       ChromeOptions co = new ChromeOptions();
	       co.addArguments("--disable-infobars");
	       co.addArguments("--disable-notifications");
	       driver = new ChromeDriver(co);
	       driver.manage().window().maximize();
	       }
		
   }
	
	//called after the test method
	@AfterMethod
	//use to close the browser
	public void quitBrowser() {
		  
		driver.quit();
	}
	
	
	//Hotel Rating Slider
	public void rslider()
	{
		Actions action2 = new Actions(driver);
		WebElement Slider2 = driver.findElement(By.xpath(prop.getProperty("rateslider")));
	    action2.dragAndDropBy(Slider2, -50,0).build().perform();
	}
	
	
	//Hotel Max Price Range Slider(It is not selecting 3000 value of the max slider that is why 4000)
	public void maxslider()
	{
		Actions action = new Actions(driver);
	    WebElement Slider = driver.findElement(By.xpath(prop.getProperty("maximumslide")));
	    action.dragAndDropBy(Slider, -150,0).build().perform();
	}
	
	
	//Hotel Min Price Range Slider
	public void minslider()
	{
		Actions action1 = new Actions(driver);
        WebElement Slider1 = driver.findElement(By.xpath(prop.getProperty("minimumslide")));
        action1.dragAndDropBy(Slider1, 1, 0).build().perform();
	}
	
//	public boolean retryingFindClick() {
//	    boolean result = false;
//	    int attempts = 0;
//	    while(attempts < 2) {
//	        try {
//	            driver.findElement(By.xpath("//*[@id=\"hds-marquee\"]/div[2]/div/div/form/div[4]/button")).click();
//	            result = true;
//	            break;
//	        } catch(StaleElementReferenceException e) {
//	        }
//	        attempts++;
//	    }
//	    return result;
//	}
	
	//Hotel Searching by location and dates.
	public void search() {
		
		//clicking on search box
		driver.findElement(By.xpath(prop.getProperty("searchbox"))).click();
		
		//searching for mumbai location
		driver.findElement(By.xpath(prop.getProperty("locatemumbai"))).click();

		//Clicking on check in calendar
		driver.findElement(By.xpath(prop.getProperty("Incalender"))).click();
		
		//clicking on check in date
		driver.findElement(By.xpath(prop.getProperty("checkin"))).click();
		
		//clicking on check out calendar
		driver.findElement(By.xpath(prop.getProperty("Outcalender"))).click();
		
		//clicking on check out date
		driver.findElement(By.xpath(prop.getProperty("checkout"))).click();
		
		
    	//clicking on search button
		driver.findElement(By.xpath("//button[@type='submit']")).click();
	}
	
	//delay
	public void sleep(int x)
	{
		try {
			Thread.sleep(x*1000);
		} catch (InterruptedException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	//Storing hotel names and prices and printing values
	public void selectingHotels()
	{
		Hotel= driver.findElements(By.xpath(prop.getProperty("hotelname")));
	    price=driver.findElements(By.xpath(prop.getProperty("pricename")));
	    	
	    	 for(int i=0;i<5;i++)
	         {System.err.println(i+1+". "+Hotel.get(i).getText()+"  "+price.get(i).getText());}
	}
	
	@Test
	public void ExpediaHotel() throws IOException
	{
		String url = "https://in.hotels.com";
		
       //calling browser
		driver.get(url);
		
		//Giving Pause to steps
		driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
		

		//Calling search function to search hotels
		search();
		
		
		//Setting a upper limit for prices
		maxslider();
	   
	   // sleep(5);
	   
	    //Setting a lower limit for prices
	    minslider();

        //Wait of 5 seconds
        sleep(5);
        
        //calling function to listing hotels
		selectingHotels();
    
		
		//Creating Excel file 
    	 @SuppressWarnings("resource")
			XSSFWorkbook workbook = new XSSFWorkbook();
		    XSSFSheet sheet = workbook.createSheet("Details");

		    Row row = sheet.createRow(0);
		    row.createCell(0).setCellValue("SNo.");
		    row.createCell(1).setCellValue("Hotel Name");
		    row.createCell(2).setCellValue("Price");

		    for(int i=0;i<5;i++)
		    {
		    	row = sheet.createRow(i+1);
		    	row.createCell(0).setCellValue(i+1);
		    	row.createCell(1).setCellValue(Hotel.get(i).getText());
		    	row.createCell(2).setCellValue(price.get(i).getText());
		    }


		    FileOutputStream fos = new FileOutputStream(new File("Hotel.xlsx"));

		    workbook.write(fos);

		    fos.close();		
		
	}


}
