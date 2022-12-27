package Rule_list;



import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.concurrent.TimeUnit;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import Excel_Functions.Excel_Functions;

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
			 driver.get("http://qms360.sleepuat.renewbuy.in/");
			 
			 
			 WebElement uName = driver.findElement(By.xpath("//*[@id='email']"));
			 WebElement pswd =driver.findElement(By.xpath("//*[@id='exampleInputPolicy']"));
			 
			 WebElement loginBtn =
			 driver.findElement(By.xpath("//button[@type='submit']"));
			 
			 
			 uName.sendKeys("admin@renewbuy.in"); pswd.sendKeys("test");
			 Thread.sleep(2000);
			 
			 loginBtn.click(); 
			 
			 Thread.sleep(2000); 
			 WebElement admin =driver.findElement(By.xpath("//span[text()='Admin']")); 
			 admin.click();
			 Thread.sleep(2000);
			 WebElement master_list = driver.findElement(By.xpath("//span[@class='d-none f-14 d-lg-inline fontsubmenu' and contains(., 'Master List')]")); 
			 master_list.click(); 
			 Thread.sleep(35000);
			 WebElement rule_list = driver.findElement(By.xpath("//label[text()='Rule list']"));
			 rule_list.click();
			 Thread.sleep(5000);
			 
			 for (int r=1;r<=rowCount;r++)
			 {
			 WebElement createrule = driver.findElement(By.xpath("//button[@class='text-light main-btn createChannelBtn' and contains(., ' Create New Rule')]")); 
			 createrule.click(); 
			 Thread.sleep(3000);
			// WebElement element=driver.findElement(By.xpath("//table[@class='table word-warp table-space']/tbody//tr[last()]"));
			 //JavascriptExecutor js = (JavascriptExecutor)driver;
			 //js.executeScript("arguments[0].scrollIntoView();", element);
			 
			 Thread.sleep(3000);
			 
		   // MISP Name selection  
			 WebElement misp_name = driver.findElement(By.
			 xpath("//span[@class='dropdown-btn']//span[text()='Misp Name']"));
			 misp_name.click(); 
			 String mispname=sheet.getRow(r).getCell(0).getStringCellValue().toUpperCase();
			// String mispnew=mispname.toUpperCase();
			// System.out.println(mispnew);
			// WebElement misp_name_search = driver.findElement(By.xpath("//li//input[@placeholder='Search']"));
		   
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
			 String dealer_code_name=sheet.getRow(r).getCell(1).getStringCellValue().toLowerCase();
			
			 WebElement dealer_code_selected= driver.findElement(By.xpath("//div[text()='"+dealer_code_name+"']")); 
			 dealer_code_selected.click();
			System.out.println("Dealer "+dealer_code_name+" selected");
			Thread.sleep(3000);
			
	// Workshop code selected
			 
			 WebElement workshop_code = driver.findElement(By.
			 xpath("//span[@class='dropdown-btn']//span[text()='Workshop Code']"));
			 workshop_code.click(); Thread.sleep(3000);
			 String workshop_code_name=sheet.getRow(r).getCell(2).getStringCellValue().toLowerCase();
			
				
			 WebElement workshop_code_selected= driver.findElement(By.xpath("//div[text()='"+workshop_code_name+"']")); 
			 workshop_code_selected.click();
			 System.out.println("Workshop "+workshop_code_name+" selected");
			 Thread.sleep(3000);
			 
			 
	 //policy type selection 
			 WebElement policy_type1 = driver.findElement(By.xpath("//input[@class='checkboxmisp' and @name='rulePolicyType+1']"));
			 policy_type1.click(); Thread.sleep(3000);
			 System.out.println("policy_type1 selected");
			 
			 WebElement policy_type2 = driver.findElement(By.xpath("//input[@class='checkboxmisp' and @name='rulePolicyType+2']"));
			 policy_type2.click(); Thread.sleep(3000);
			 System.out.println("policy_type2 selected");
			 
	 // IC Selection
			 WebElement ic_name = driver.findElement(By.xpath("//span[text()='IC Name']"));
			 ic_name.click(); Thread.sleep(2000);
			 
			 String ic=sheet.getRow(r).getCell(3).getStringCellValue();
				
			 WebElement ic_name_selected= driver.findElement(By.xpath("//div[text()='"+ic+"']")); 
			 ic_name_selected.click();
			 System.out.println("IC "+ic+" selected");
			 Thread.sleep(3000);
			 
	 // Issue Type Selection
			 WebElement issue_type = driver.findElement(By.xpath("//span[text()='Issue Type']"));
			 issue_type.click(); Thread.sleep(2000);

			 String issue=sheet.getRow(r).getCell(4).getStringCellValue();
				
			 WebElement issue_type_selected= driver.findElement(By.xpath("//div[text()='"+issue+"']")); 
			 issue_type_selected.click();
			 System.out.println("Issue_Type "+issue+" selected");
			 Thread.sleep(3000);

	 // Issue sub-Type Selection
			 WebElement issue_subtype = driver.findElement(By.xpath("//span[text()='Issue Sub Type']"));
			 issue_subtype.click(); Thread.sleep(2000);

			 String sub_issue=sheet.getRow(r).getCell(5).getStringCellValue();
				
			 WebElement issue_subtype_selected= driver.findElement(By.xpath("//div[text()='"+sub_issue+"']")); 
			 issue_subtype_selected.click();
			 System.out.println("Issue sub-Type "+sub_issue+" selected");
			 Thread.sleep(3000);

	// Channel Mappings		 
			 WebElement channel1 = driver.findElement(By.xpath("//input[@name='createL0MappingValue']"));
			 channel1.click(); Thread.sleep(3000);
			
	// TAT Format
			 WebElement tat_format = driver.findElement(By.xpath("//input[@id='radioTATL1-Create2']"));
			 tat_format.click(); Thread.sleep(3000);

			 
	//TAT 1 selection
			 WebElement tat_1 = driver.findElement(By.xpath("//input[@name='TATL1RuleDays']"));
			 tat_1.click(); Thread.sleep(2000);
			 
			 Double tat1=sheet.getRow(r).getCell(10).getNumericCellValue();
			 int value1 = (int)Math.round(tat1);
			
			 tat_1.sendKeys(String.valueOf(value1));
			 System.out.println("Entered TAT-1 value is "+String.valueOf(value1)+ ".");
			 Thread.sleep(3000);
			 

    //TAT 2 selection
			 WebElement tat_2 = driver.findElement(By.xpath("//input[@name='TATL2RuleDays']"));
			 tat_2.click(); Thread.sleep(2000);
			 
			 Double tat2=sheet.getRow(r).getCell(11).getNumericCellValue();
			 int value2 = (int)Math.round(tat2);
			
			 tat_2.sendKeys(String.valueOf(value2));
			 System.out.println("Entered TAT-1 value is "+String.valueOf(value2)+ ".");
			 Thread.sleep(3000);

    //TAT 3 selection
			 WebElement tat_3 = driver.findElement(By.xpath("//input[@name='TATL3RuleDays']"));
			 tat_3.click(); Thread.sleep(2000);
			 Double tat3=sheet.getRow(r).getCell(12).getNumericCellValue();
			 int value3 = (int)Math.round(tat3);
			
			 tat_3.sendKeys(String.valueOf(value3));
			 System.out.println("Entered TAT-1 value is "+String.valueOf(value3)+ ".");
			 Thread.sleep(3000);
			 
	// Save Button click 
			 WebElement save_rule = driver.findElement(By.xpath("//button[@class='saveBtnChannel misptableAnchorTagSave']"));
			 save_rule.click(); Thread.sleep(3000);
			 System.out.println("Save Clicked");
			 

    // Sheet status update 
	
			// WebElement rule_success = driver.findElement(By.xpath("//div[text()='Rule Created Successfully']"));
			 WebElement rule_exist = driver.findElement(By.xpath("//div[text()='Rule already exist for the given combination']"));
			 //create a new cell in the row at index 6
			 XSSFCell cell = sheet.getRow(r).createCell(13);
			 
			 //check if confirmation message is displayed
			 if (rule_exist.isDisplayed()) {
				 // if the message is displayed , write PASS in the excel sheet
				 cell.setCellValue("Rule Already Exist");
				 
			 }    //else if(rule_exist.isDisplayed()) {
				  //if the message is not displayed , write FAIL in the excel sheet
				 // cell.setCellValue("Already Exist");
			     //}
			 else{

				cell.setCellValue("FAILED");
			 }
			 // Write the data back in the Excel file
			 FileOutputStream outputStream = new FileOutputStream(excelFilePath);
			 workbook.write(outputStream);
			   
			 
			  
			
			 
			 }
			 
			 
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
