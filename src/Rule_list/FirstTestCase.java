package Rule_list;



import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.concurrent.TimeUnit;

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
			 
			 File excel=new File(excelFilePath); FileInputStream inputstream=new
			 FileInputStream(excelFilePath);
			 
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
			 Thread.sleep(3000);
			 
			 for (int r=1;r<=rowCount;r++)
			 {
			 WebElement createrule = driver.findElement(By.xpath("//button[@class='text-light main-btn createChannelBtn' and contains(., ' Create New Rule')]")); 
			 createrule.click(); 
			 Thread.sleep(3000);
			 WebElement element=driver.findElement(By.xpath("//table[@class='table word-warp table-space']/tbody//tr[last()]"));
			 JavascriptExecutor js = (JavascriptExecutor)driver;
			 js.executeScript("arguments[0].scrollIntoView();", element);
			 
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
			 
			 WebElement issue_type = driver.findElement(By.xpath("//span[text()='Issue Type']"));
			 issue_type.click(); Thread.sleep(3000);
					 
			 WebElement issue_subtype = driver.findElement(By.xpath("//span[text()='Issue Sub Type']"));
			 issue_subtype.click(); Thread.sleep(3000);
					 
			 WebElement channel1 = driver.findElement(By.xpath("//input[@name='createL0MappingValue']"));
			 channel1.click(); Thread.sleep(3000);
			
			
			 WebElement tat_format = driver.findElement(By.xpath("//input[@id='radioTATL1-Create2']"));
			 tat_format.click(); Thread.sleep(3000);
			 
			 WebElement tat_1 = driver.findElement(By.xpath("//input[@name='TATL1RuleDays']"));
			 tat_1.click(); Thread.sleep(3000);
			 
			 WebElement tat_2 = driver.findElement(By.xpath("//input[@name='TATL2RuleDays']"));
			 tat_2.click(); Thread.sleep(3000);
			 
			 WebElement tat_3 = driver.findElement(By.xpath("//input[@name='TATL3RuleDays']"));
			 tat_3.click(); Thread.sleep(3000);
			 
			 WebElement save_rule = driver.findElement(By.xpath("//button[@class='saveBtnChannel misptableAnchorTagSave']"));
			 save_rule.click(); Thread.sleep(3000);
			 
			   
			 
			  
			
			 
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
