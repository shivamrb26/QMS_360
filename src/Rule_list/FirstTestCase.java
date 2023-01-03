package Rule_list;



import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.time.LocalDateTime;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.concurrent.TimeUnit;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.StaleElementReferenceException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import Excel_Functions.Excel_Functions;
import io.netty.handler.codec.marshalling.ThreadLocalUnmarshallerProvider;

public class FirstTestCase {
	

	public static void main(String[] args) throws InterruptedException, IOException,NullPointerException {
		String excelFilePath="/home/shivamsharma/Downloads/Rule Creation1.xls";
	   
		   
		   
			 // excelinit
			 
			FileInputStream inputstream=new FileInputStream(excelFilePath);
			 
			 XSSFWorkbook workbook = new XSSFWorkbook(inputstream); // XSSFSheet
			 XSSFSheet sheet=workbook.getSheet("sheet1"); //Providing sheet name XSSFSheet
			 sheet=workbook.getSheetAt(0);
			
			 int rowCount=sheet.getLastRowNum()-sheet.getFirstRowNum();
			 int cols=sheet.getRow(0).getLastCellNum();
				
				
			   
				   /*
					* for (int r=0;r<=rowCount;r++) { XSSFRow row=sheet.getRow(r); //focussed on
					* current row
					* 
					* for(int c=0;c<cols;c++) { XSSFCell cell=row.getCell(c);
					* 
					* switch(cell.getCellType()) { case STRING:
					* System.out.print(cell.getStringCellValue() +"    " ); break; case NUMERIC:
					* System.out.print(cell.getNumericCellValue()+"    "); break; case BOOLEAN:
					* System.out.print(cell.getBooleanCellValue()+"    "); break; default: break; }
					* } System.out.println(); }
					*/
			  
		   
			 System.setProperty("webdriver.chrome.driver","/home/shivamsharma/Downloads/chromedriver"); 
	      	 WebDriver driver = new ChromeDriver();
			 driver.manage().window().maximize();
			 driver.get("https://www.renewbuy.com/qms360/");
			 
			 
			 WebElement uName = driver.findElement(By.xpath("//*[@id='email']"));
			 WebElement pswd =driver.findElement(By.xpath("//*[@id='exampleInputPolicy']"));
			 
			 WebElement loginBtn =
			 driver.findElement(By.xpath("//button[@type='submit' and contains(., 'Sign In')]"));
			 
			 
			 uName.sendKeys("qmsadmin@renewbuy.com"); pswd.sendKeys("qmsadmin@2468");
			 //uName.sendKeys("admin@renewbuy.in"); pswd.sendKeys("test");
			 Thread.sleep(2000);
			 
			 loginBtn.click(); 
			 
			 Thread.sleep(2000); 
			 WebElement admin =driver.findElement(By.xpath("//span[text()='Admin']")); 
			 admin.click();
			 Thread.sleep(2000);
			 WebElement master_list = driver.findElement(By.xpath("//span[@class='d-none f-14 d-lg-inline fontsubmenu' and contains(., 'Master List')]")); 
			 master_list.click(); 
			 Thread.sleep(60000);
			 WebElement rule_list = driver.findElement(By.xpath("//label[text()='Rule list']"));
			 rule_list.click();
			 
			 List<WebElement> elm;
			 elm = driver.findElements(By.xpath("//*[@class='ngx-background-spinner bottom-right loading-background']"));
			 driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(10));
			 while (elm.size() > 0) {
				 elm = driver.findElements(By.xpath("//*[@class='ngx-background-spinner bottom-right loading-background']"));
			 }
			 driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(45));
			 System.out.println("while end");
			 Thread.sleep(3000);
			
			 WebElement ops_rule=driver.findElement(By.xpath("//button[@id='opsId']"));
			 ops_rule.click();
			 Thread.sleep(4000);
			 
			 for (int r=1;r<=rowCount;r++)
			 {
            
			List<WebElement> elm1;
        elm1 = driver.findElements(By.xpath("//*[@class='ngx-background-spinner bottom-right loading-background']"));
        driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(10));
        while (elm1.size() > 0) {
            elm1 = driver.findElements(By.xpath("//*[@class='ngx-background-spinner bottom-right loading-background']"));
        }
        driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(15));
        System.out.println("while end");
		// Thread.sleep(15000);
            WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(70));
			wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//button[@class='text-light main-btn createChannelBtn' and contains(., ' Create New Rule')]"))).click();
            //Boolean visible = wait.ignoring(StaleElementReferenceException.class).until(ExpectedConditions.and(ExpectedConditions.visibilityOfElementLocated(By.xpath("//button[@class='text-light main-btn createChannelBtn' and contains(., ' Create New Rule')]")),
			//	ExpectedConditions.elementToBeClickable(By.xpath("//button[@class='text-light main-btn createChannelBtn' and contains(., ' Create New Rule')]")),
			//	ExpectedConditions.presenceOfElementLocated(By.xpath("//button[@class='text-light main-btn createChannelBtn' and contains(., ' Create New Rule')]"))));

			// WebElement createrule = driver.findElement(By.xpath("//button[@class='text-light main-btn createChannelBtn' and contains(., ' Create New Rule')]")); 
			// WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(15));
		     //WebElement element = wait.until(ExpectedConditions.elementToBeClickable(createrule));
			
			 
			// WebElement element=driver.findElement(By.xpath("//table[@class='table word-warp table-space']/tbody//tr[last()]"));
			 //JavascriptExecutor js = (JavascriptExecutor)driver;
			 //js.executeScript("arguments[0].scrollIntoView();", element);
			 
			  Thread.sleep(1000);
			 
		  // MISP Name selection  
			 WebElement misp_name = driver.findElement(By.
			 xpath("//span[@class='dropdown-btn']//span[text()='Misp Name']"));
			 misp_name.click(); 
			 String mispname=sheet.getRow(r).getCell(0).getStringCellValue();
			 if (mispname.equals("All")){
				mispname="Select All";
			 }
			// String mispnew=mispname.toUpperCase();
			// System.out.println(mispnew);
			WebElement misp_name_search = driver.findElement(By.xpath("//li//input[@placeholder='Search']"));
		   
	   //	  misp_name_search.click();
			// misp_name_search.sendKeys(mispname);
			 WebElement misp_name_searched = driver.findElement(By.xpath("//div[text()='"+mispname+"']")); 
			 misp_name_searched.click();
			 //misp_name_search.sendKeys(sheet.getRow(r).getCell(0).getStringCellValue() +",");
			 System.out.println("MISP "+mispname+" selected");
			 Thread.sleep(3000);
			
   // Dealer code selection
			 
			 WebElement dealer_code = driver.findElement(By.
			 xpath("//span[@class='dropdown-btn']//span[text()='Dealer Code']"));
			 dealer_code.click(); 
			 Thread.sleep(2000);
			 String dealer_code_name=sheet.getRow(r).getCell(1).getStringCellValue();
			 if (dealer_code_name.equals("All")){
				dealer_code_name="Select All";
			 }
			 WebElement dealer_code_selected= driver.findElement(By.xpath("//div[text()='"+dealer_code_name+"']")); 
			 dealer_code_selected.click();
			System.out.println("Dealer "+dealer_code_name+" selected");
			Thread.sleep(4000);
			
	// Workshop code selected
	        
			//  WebElement workshop_code = driver.findElement(By.xpath("//span[@class='dropdown-btn']//span[text()='Workshop Code']"));
			//  workshop_code.click(); Thread.sleep(3000);
			//  String workshop_code_name=sheet.getRow(r).getCell(2).getStringCellValue();
			// if(workshop_code_name.isEmpty()){

			// 	workshop_code.click(); Thread.sleep(1000);

			// }
			// else{
			//  WebElement workshop_code_selected= driver.findElement(By.xpath("//div[text()='"+workshop_code_name+"']")); 
			//  workshop_code_selected.click();
			//  System.out.println("Workshop "+workshop_code_name+" selected");
			//  Thread.sleep(3000);
			 

			// 	}
			 
			 
	 //policy type selection 
			 WebElement policy_type1 = driver.findElement(By.xpath("//input[@class='checkboxmisp' and @name='rulePolicyType+1']"));
			 policy_type1.click(); 
			 System.out.println("policy_type1 selected");
			 Thread.sleep(2000);
			 WebElement policy_type2 = driver.findElement(By.xpath("//input[@class='checkboxmisp' and @name='rulePolicyType+2']"));
			 policy_type2.click(); 
			 System.out.println("policy_type2 selected");Thread.sleep(1500);
			 
	 // IC Selection
			 WebElement ic_name = driver.findElement(By.xpath("//span[text()='IC Name']"));
			 ic_name.click(); 
			 
			 String ic=sheet.getRow(r).getCell(3).getStringCellValue();
				
			 WebElement ic_name_selected= driver.findElement(By.xpath("//div[text()='"+ic+"']")); 
			 ic_name_selected.click();
			 System.out.println("IC "+ic+" selected");
			 Thread.sleep(500);
			 
	 // Issue Type Selection
			 WebElement issue_type = driver.findElement(By.xpath("//span[text()='Issue Type']"));
			 issue_type.click(); Thread.sleep(100);

			 String issue=sheet.getRow(r).getCell(4).getStringCellValue();
				
			 WebElement issue_type_selected= driver.findElement(By.xpath("//div[text()='"+issue+"']")); 
			 issue_type_selected.click();
			 System.out.println("Issue_Type "+issue+" selected");
			 Thread.sleep(3000);

	 // Issue sub-Type Selection
			 WebElement issue_subtype = driver.findElement(By.xpath("//span[text()='Issue Sub Type']"));
			 issue_subtype.click(); Thread.sleep(3000);

			 String sub_issue=sheet.getRow(r).getCell(5).getStringCellValue();
				
			 WebElement issue_subtype_selected= driver.findElement(By.xpath("//div[text()='"+sub_issue+"']")); 
			 issue_subtype_selected.click();
			 System.out.println("Issue sub-Type "+sub_issue+" selected");
			 Thread.sleep(3000);

	// Channel Mappings		 
			 WebElement l0 = driver.findElement(By.xpath("//input[@name='createL0MappingValue']"));
			 l0.click(); Thread.sleep(3000);
			 String l0_channel =sheet.getRow(r).getCell(6).getStringCellValue();
			 
			 WebElement l0_selected= driver.findElement(By.xpath("//button[@class='l0maiingbtn' and contains(., '"+l0_channel+"')]")); 
			 l0_selected.click();
			 System.out.println("L0 "+l0_channel+" selected");
			 Thread.sleep(3000);

			 WebElement l1 = driver.findElement(By.xpath("//input[@name='createL1MappingValue']"));
			 l1.click(); Thread.sleep(3000);
			 String l1_channel=sheet.getRow(r).getCell(7).getStringCellValue();
			 WebElement l1_selected= driver.findElement(By.xpath("//button[@class='l0maiingbtn' and contains(., '"+l1_channel+"')]")); 
			 l1_selected.click();
			 System.out.println("L1 "+l1_channel+" selected");
			 Thread.sleep(3000);

			 WebElement l2 = driver.findElement(By.xpath("//input[@name='createL2MappingValue']"));
			 l2.click(); Thread.sleep(3000);
			 String l2_channel=sheet.getRow(r).getCell(8).getStringCellValue();
			 WebElement l2_selected= driver.findElement(By.xpath("//button[@class='l0maiingbtn' and contains(., '"+l2_channel+"')]"));  
			 l2_selected.click();
			 System.out.println("L2 "+l2_channel+" selected");
			 Thread.sleep(3000);
           
             WebElement l3 = driver.findElement(By.xpath("//input[@name='createL3MappingValue']"));
			 l3.click(); Thread.sleep(3000);
			 String l3_channel=sheet.getRow(r).getCell(9).getStringCellValue();
			 WebElement l3_selected= driver.findElement(By.xpath("//button[@class='l0maiingbtn' and contains(., '"+l3_channel+"')]"));  
			 l3_selected.click();
			 l3.click();
			 System.out.println("L3 "+l3_channel+" selected");
			 Thread.sleep(3000);
			 
	//TAT 1 selection
			
			Date tat1hrs=sheet.getRow(r).getCell(10).getDateCellValue();
			if(tat1hrs != null){
			WebElement tat_1hrs = driver.findElement(By.xpath("//input[@name='TATL1Rule']"));
			  tat_1hrs.click(); Thread.sleep(2000);
			// Date tat1=sheet.getRow(r).getCell(10).getDateCellValue();
			SimpleDateFormat formatTime = new SimpleDateFormat("HH:mm");
			String timeStamp_tat1 =formatTime.format(tat1hrs);
			tat_1hrs.sendKeys(String.valueOf(timeStamp_tat1));
            System.out.println("Entered TAT-1(hrs) value is "+String.valueOf(timeStamp_tat1)+ ".");
			Thread.sleep(3000);
		}
		else{
					
	// TAT Format
	WebElement tat_format = driver.findElement(By.xpath("//input[@id='radioTATL1-Create2']"));
	tat_format.click(); Thread.sleep(3000);  

             WebElement days_path = driver.findElement(By.xpath("//input[@name='TATL1RuleDays']"));
			 
			 days_path.click(); Thread.sleep(2000);
			 
			 Double tat1days=sheet.getRow(r).getCell(11).getNumericCellValue();
			  int value_tat1 = (int)Math.round(tat1days);
			
			
			 days_path.sendKeys(String.valueOf(value_tat1));
			 System.out.println("Entered TAT-1(Days) value is "+String.valueOf(value_tat1)+ ".");
			 Thread.sleep(3000);




		}

		//TAT 2 selection
			
		Date tat2hrs=sheet.getRow(r).getCell(12).getDateCellValue();
		if(tat2hrs != null){
		WebElement tat_2hrs = driver.findElement(By.xpath("//input[@name='TATL2Rule']"));
		  tat_2hrs.click(); Thread.sleep(2000);
		// Date tat1=sheet.getRow(r).getCell(10).getDateCellValue();
		SimpleDateFormat formatTime = new SimpleDateFormat("HH:mm");
		String timeStamp_tat2 =formatTime.format(tat2hrs);
		tat_2hrs.sendKeys(String.valueOf(timeStamp_tat2));
		System.out.println("Entered TAT-1(hrs) value is "+String.valueOf(timeStamp_tat2)+ ".");
		Thread.sleep(3000);
	}
	else{
				
// TAT Format
//WebElement tat_format = driver.findElement(By.xpath("//input[@id='radioTATL1-Create2']"));
//tat_format.click(); Thread.sleep(3000);  

		 WebElement days_path_tat2 = driver.findElement(By.xpath("//input[@name='TATL2RuleDays']"));
		 
		 days_path_tat2.click(); Thread.sleep(2000);
		 
		 Double tat2days=sheet.getRow(r).getCell(13).getNumericCellValue();
		  int value_tat2 = (int)Math.round(tat2days);
		
		
		 days_path_tat2.sendKeys(String.valueOf(value_tat2));
		 System.out.println("Entered TAT-1(Days) value is "+String.valueOf(value_tat2)+ ".");
		 Thread.sleep(3000);




	}

//     //TAT 2 selection
// 			//  WebElement tat_2 = driver.findElement(By.xpath("//TAT 2 selection
			
// 		Date tat2hrs=sheet.getRow(r).getCell(14).getDateCellValue();
// 		if(tat2hrs != null){
// 		WebElement tat_2hrs = driver.findElement(By.xpath("//input[@name='TATL2Rule']"));
// 		  tat_2hrs.click(); Thread.sleep(2000);
// 		// Date tat1=sheet.getRow(r).getCell(10).getDateCellValue();
// 		SimpleDateFormat formatTime = new SimpleDateFormat("HH:mm");
// 		String timeStamp_tat2 =formatTime.format(tat2hrs);
// 		tat_2hrs.sendKeys(String.valueOf(timeStamp_tat2));
// 		System.out.println("Entered TAT-1(hrs) value is "+String.valueOf(timeStamp_tat2)+ ".");
// 		Thread.sleep(3000);
// 	}
// 	else{
				
// // TAT Format
// //WebElement tat_format = driver.findElement(By.xpath("//input[@id='radioTATL1-Create2']"));
// //tat_format.click(); Thread.sleep(3000);  

// 		 WebElement days_path_tat2 = driver.findElement(By.xpath("//input[@name='TATL2RuleDays']"));
		 
// 		 days_path_tat2.click(); Thread.sleep(2000);
		 
// 		 Double tat2days=sheet.getRow(r).getCell(15).getNumericCellValue();
// 		  int value_tat2 = (int)Math.round(tat2days);
		
		
// 		 days_path_tat2.sendKeys(String.valueOf(value_tat2));
// 		 System.out.println("Entered TAT-1(Days) value is "+String.valueOf(value_tat2)+ ".");
// 		 Thread.sleep(3000);




// 	}//input[@name='TATL2RuleDays']"));
			//  tat_2.click(); Thread.sleep(2000);
			 
			//  Double tat2=sheet.getRow(r).getCell(11).getNumericCellValue();
			//  int value2 = (int)Math.round(tat2);
			
			//  tat_2.sendKeys(String.valueOf(value2));
			//  System.out.println("Entered TAT-1 value is "+String.valueOf(value2)+ ".");
			//  Thread.sleep(3000);


			//TAT 3 selection
			
		Date tat3hrs=sheet.getRow(r).getCell(14).getDateCellValue();
		if(tat3hrs != null){
		WebElement tat_3hrs = driver.findElement(By.xpath("//input[@name='TATL3RuleDays']"));
		  tat_3hrs.click(); Thread.sleep(2000);
		// Date tat1=sheet.getRow(r).getCell(10).getDateCellValue();
		SimpleDateFormat formatTime = new SimpleDateFormat("HH:mm");
		String timeStamp_tat3 =formatTime.format(tat3hrs);
		tat_3hrs.sendKeys(String.valueOf(timeStamp_tat3));
		System.out.println("Entered TAT-1(hrs) value is "+String.valueOf(timeStamp_tat3)+ ".");
		Thread.sleep(3000);
	}
	else{
				
// TAT Format
//WebElement tat_format = driver.findElement(By.xpath("//input[@id='radioTATL1-Create2']"));
//tat_format.click(); Thread.sleep(3000);  

		 WebElement days_path_tat3 = driver.findElement(By.xpath("//input[@name='TATL3RuleDays']"));
		 
		 days_path_tat3.click(); Thread.sleep(2000);
		 
		 Double tat3days=sheet.getRow(r).getCell(15).getNumericCellValue();
		  int value_tat3 = (int)Math.round(tat3days);
		
		
		 days_path_tat3.sendKeys(String.valueOf(value_tat3));
		 System.out.println("Entered TAT-1(Days) value is "+String.valueOf(value_tat3)+ ".");
		 Thread.sleep(3000);




	}

    //TAT 3 selection
			//  WebElement tat_3 = driver.findElement(By.xpath("//input[@name='TATL3RuleDays']"));
			//  tat_3.click(); Thread.sleep(2000);
			//  Double tat3=sheet.getRow(r).getCell(12).getNumericCellValue();
			//  int value3 = (int)Math.round(tat3);
			
			//  tat_3.sendKeys(String.valueOf(value3));
			//  System.out.println("Entered TAT-1 value is "+String.valueOf(value3)+ ".");
			//  Thread.sleep(3000);
			 
	// // Save Button click 
			 WebElement save_rule = driver.findElement(By.xpath("//button[@class='saveBtnChannel misptableAnchorTagSave']"));
			 save_rule.click(); 
			 System.out.println("Save Clicked");
			 Thread.sleep(1000);
			 
			 

    // Sheet status update 

			 String message = driver.findElement(By.xpath("(//*[@class='alert translate-bottom alert-success'])//div[2]")).getText();
             //System.out.println(rule_success);
			//WebElement rule_exist = driver.findElement(By.xpath("//div[text()='Rule already exist for the given combination']"));
			 //create a new cell in the row 
			 XSSFCell cell = sheet.getRow(r).createCell(16);
			
     
			 
			 //check if confirmation message is displayed
			 if (message.equals("Rule successfully created")) {
				 // if the message is displayed , write PASS in the excel sheet
				 cell.setCellValue("Rule created");
				 System.out.println("created rule");
				 
			 }   
			  else if(message.equals("Rule already exist for the given combination")) {
				 
				 cell.setCellValue("Already Exist");
				 System.out.println("rule exist");
			     }
			 
			 else{
				cell.setCellValue("Failed");
				System.out.println("Failed");

			}
			 // Write the data back in the Excel file
			 FileOutputStream outputStream = new FileOutputStream(excelFilePath);
			 workbook.write(outputStream);
			   
			 Thread.sleep(4000);
			  
			
			 
			 }
			 driver.quit();
			 
	}}
















/*try {

//Locating web element
WebElement logoutBtn = driver.findElement(By.xpath("//div[@class='text-right col-md-5 col-sm-12']//button[@id='submit']"));
//Validating presence of element				
if(logoutBtn.isDisplayed()){
	
	//Performing action on web element
	logoutBtn.click();
	System.out.println("LogOut Successful!");
}
} 
catch (Exception e) {
System.out.println("Incorrect login....");
}
*/	
//Closing browser session
