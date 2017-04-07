
// Title: 		AutoFox
// Purpose:		Automates Firefox for automated data entry
// Dependencies:	Firefox WebDriver, JUnit, Selenium
// Disclaimer:		Not plug-n-play. Script was built for a specific task on specific website.
// Author:		Brian White
// Date:		July, 2014

import java.awt.Toolkit;
import java.io.FileInputStream;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.concurrent.TimeUnit;
import jxl.Sheet;
import jxl.Workbook;
import org.junit.*;
import org.openqa.selenium.*;
import org.openqa.selenium.firefox.FirefoxDriver;

public class mb5 {
  private WebDriver driver;
  private String baseUrl;
  private String newLine = System.getProperty("line.separator");
  private String FilePath;
  private Sheet sh;
  private Integer row;
  private JavascriptExecutor js;
  
  // Initialize JUnit BEFORE
  @Before
  public void setUp() throws Exception {
    driver = new FirefoxDriver();
    if (driver instanceof JavascriptExecutor) {
        js = (JavascriptExecutor)driver;
    }
    baseUrl = "https://www.somefakesite.com/";
    driver.manage().timeouts().implicitlyWait(3, TimeUnit.SECONDS);
    
    // Set the path of our input file
    FilePath = "C://someinputfile.xls";
    
    FileInputStream fs = new FileInputStream(FilePath);
    Workbook wb = Workbook.getWorkbook(fs);	
    sh = wb.getSheet(0);
  }
 
  // Run our JUnit Test
  @Test
  public void mbLoop() throws Exception {
	  
	  // Set # of iterations (EOF)
	  for(row=1; row<sh.getRows(); row++){
		  
		    String sMaster = sh.getCell(0,row).getContents();
			String sMsg = sh.getCell(1,row).getContents();
			try{
				driver.get(baseUrl + "somepartialURL");
				LoginCode(sMaster, "somepw");
				SelectShipTo(sMaster);
				EnterMB(sMaster, sMsg);
				driver.manage().deleteAllCookies();
	  
			}catch (UnhandledAlertException a){
				driver.navigate().refresh();
				System.out.print(sMaster + "- Unhandled Alert Exception Encountered");
		  
			}catch (TimeoutException b) {
				driver.navigate().refresh();
				System.out.print(sMaster + "- Page Timed Out");
			}
	  }
  }
  
  // Part I: Login to the site, handle any errors and redirect appropriately
  public void LoginCode (String strMaster, String strPassword) throws Exception{
		
		try{
		driver.findElement(By.id("SomeWebElementID")).clear();
		driver.findElement(By.id("SomeWebElementID")).sendKeys(strMaster);
		driver.findElement(By.id("SomeWebElementID")).clear();
		driver.findElement(By.id("SomeWebElementID")).sendKeys("SALESONLY");
		driver.findElement(By.id("SomeWebElementID1")).clear();
		driver.findElement(By.id("SomeWebElementID")).sendKeys(strPassword);
		driver.findElement(By.cssSelector("a.btn_action > span.left")).click();
		}catch (NoSuchElementException a){
			driver.manage().deleteAllCookies();
			System.out.print(strMaster + "- No Login Element Found - FYI" + newLine);
			driver.get(baseUrl + "somepartialURL");
			LoginCode(strMaster, strPassword);
		}
	}
  
  // Part II: Clicks through screening process post-login.  Handles errors and redirects appropriately.
  public void SelectShipTo(String strMaster) throws Exception{
	  String errMsg = "";
	  
	  try{
		  errMsg = driver.findElement(By.id("loginErrorAlertParagraph")).getText();
		
		// Trigger JS if element fails to load.  78% success rate.
	  } catch(NoSuchElementException n){
		  try{
			  js.executeScript("javascript:document.ShipToForm.LaunchItems.value='custmsgs';document.ShipToForm.LaunchItems.name='URL';doOK(document.ShipToForm)");
		 
		   // Log error in console if JS doesn't resolve
		  }catch (Exception b){
			  System.out.print(strMaster + "- Login Error Save Likely Fail");
		  }
	  }
	  
	  if (errMsg.equals("Please re-enter your password and click \"Log in\" to begin a new visit.")){
		  PostErrorLogin(strMaster, "somepw");
	  }
	  
	  if (errMsg.equals("You have entered an invalid Company ID, User ID or password. The password field is case sensitive. Please try again.")){
		  PostErrorLogin(strMaster, "somepw");
	  }
	  
	  if (errMsg.equals("You have entered an invalid Customer ID, User ID or password. Please try again.")){
		  PostErrorLogin(strMaster, "somepw");
	  }
	  
	  // If the login reflects an inactive account, a special error is logged in the console
	  if (errMsg.equals("We're sorry; this Customer ID or User ID is no longer active in our system. Please contact your Administrator.")){
		  System.out.print(strMaster + "- Inactive Account");
	  }
	  
	  if (errMsg.equals("You did not enter a value into the Customer ID field. This is a required field. Please enter it now.")){
		  driver.manage().deleteAllCookies();
		  driver.get(baseUrl + "somepartialURL");
		  LoginCode(strMaster, "somepw");
		  SelectShipTo(strMaster);
	  }
	}
  
  // Part III: Enters then saves data.  Uses site response to confirm if save is successful.
  public void EnterMB(String strMaster, String strMsg) throws Exception{
	  
	  try{
		  	// Sleeps the script for 2 seconds.  Important when site is experiencing high-traffic and page load time is affected
			Thread.sleep(2000);
			driver.findElement(By.id("APPLY_TO_ALL_BTS")).click();
	
			driver.findElement(By.id("BT_MESSAGE")).clear();
			driver.findElement(By.id("BT_MESSAGE")).sendKeys(strMsg);
			
			driver.findElement(By.cssSelector("a.btn_action > span.left")).click();
			Assert.assertEquals("Messages have been successfully updated.", driver.findElement(By.cssSelector("p.orangeTextBold")).getText());
			System.out.print(strMaster + "- Save Passed" + newLine);
		} catch (Exception a){
			System.out.print(strMaster + "- Save Failed" + newLine);
		}
			
	}

  // Part I.5 Handles Login Errors, redirects appropriately
  public void PostErrorLogin(String strMaster, String strPassword) throws Exception{
	  
	  try{
	  driver.findElement(By.id("SomeWebElement")).sendKeys(strPassword);
	  
	  // Triggers JS for log out
	  js.executeScript("javascript:subDoLogonError()");
	  Thread.sleep(1500);
	  SelectShipTo(strMaster);
	  } catch (Exception e){
		  System.out.print(strMaster + "- Post Error Login Fail");
	  }
  }

  // After all iterations have been completed, done() is executed
@After
  public void done() throws Exception {
	
	 // Logs completion time in console
	 String timeStamp = new SimpleDateFormat("yyyyMMdd_HHmmss").format(Calendar.getInstance().getTime());
	 System.out.print("Completed on " + timeStamp);
	 
	 // Plays a really loud, disruptive sound that shatters my ear drums, notifying me the script has finished running
	 final Runnable runnable = (Runnable) Toolkit.getDefaultToolkit().getDesktopProperty("win.sound.exclamation");
	 if (runnable != null) runnable.run();
	 
	 driver.quit();
	  
}
  
}
  
