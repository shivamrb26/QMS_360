package Excel_Functions;


	import java.io.File;
	import java.io.FileInputStream;
	import java.io.FileNotFoundException;
	import java.io.FileOutputStream;
	import java.io.IOException;

	import org.apache.poi.xssf.usermodel.XSSFCell;
	import org.apache.poi.xssf.usermodel.XSSFRow;
	import org.apache.poi.xssf.usermodel.XSSFSheet;
	import org.apache.poi.xssf.usermodel.XSSFWorkbook;

	public class Excel_Functions {
		
	   private static XSSFWorkbook workbook;
	    private static XSSFSheet sheet;
	    private static XSSFRow row;
	    private static XSSFCell cell;

	 
	   
	    public void  excelInit(String excelFilePath ) throws IOException {
	    	
	    	File excel=new File(excelFilePath);
			 FileInputStream inputstream=new FileInputStream(excel);
			
			 workbook = new XSSFWorkbook(inputstream);
				// XSSFSheet sheet=workbook.getSheet("sheet1");   //Providing sheet name
			 sheet=workbook.getSheetAt(0);
	    }

	    
	    
	    public String getData(int rowNumber,int cellNumber) { 
	    cell=sheet.getRow(rowNumber).getCell(cellNumber);
	    return cell.getStringCellValue();
	    }
	    
	    
	    
	    
	    public int getRowCount() { 
	    int rowCount=sheet.getLastRowNum()-sheet.getFirstRowNum();
	    return rowCount;
	    }
	    
	    public int getColCount() { 
	        int colCount=sheet.getRow(0).getLastCellNum();
	        return colCount;
	        }
	    public void setCellValue(int rowNum,int cellNum,String cellValue,String excelFilePath) throws IOException {
	    	    
	    	sheet.getRow(rowNum).createCell(cellNum).setCellValue(cellValue);
	        
	    	FileOutputStream outputStream = new FileOutputStream(excelFilePath);
	    	workbook.write(outputStream);
	    	workbook.close();
}}
