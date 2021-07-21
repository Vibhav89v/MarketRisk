package generics;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import common.AutomationConstants;


public class ExcelNew implements AutomationConstants{

		public FileInputStream fis=null;
		public FileOutputStream fos=null;
		public XSSFWorkbook workbook=null;
		public XSSFSheet sheet =null;
		public XSSFRow row=null;
		public XSSFCell cell=null;
		String xlFilePath;
		public ExcelNew(String xlFilePath) throws Exception{
			this.xlFilePath= xlFilePath;
			fis= new FileInputStream(xlFilePath);
			workbook=new XSSFWorkbook(fis);
			fis.close();
		}
		
		public  String getCellData(String sheetName,  int rowNum, String colName) {
			try {
				int colNum= -1;
			  sheet=workbook.getSheet(sheetName);
			  row=sheet.getRow(rowNum);
			  for(int i=0; i<row.getLastCellNum(); i++) {
				  if(row.getCell(i).getStringCellValue().trim().equals(colName))
					  colNum=i;
			  }
			  row= sheet.getRow(rowNum -1);
			  cell = row.getCell(colNum);
			 
			  
			  if(cell.getCellTypeEnum()==CellType.STRING) {
				  return cell.getStringCellValue();
			  }
			  else if(cell.getCellTypeEnum()==CellType.NUMERIC || cell.getCellTypeEnum() == CellType.FORMULA) {
				  String cellValue=String.valueOf(cell.getNumericCellValue());
				  if(HSSFDateUtil.isCellDateFormatted(cell)) {
					  DateFormat dt= new SimpleDateFormat("dd-mmm-yy");
					  Date date=cell.getDateCellValue();
					  cellValue=dt.format(date);
				  }
				  return cellValue;
			  }
			  else if(cell.getCellTypeEnum()==CellType.BLANK) {
				  return "";
			  }
			  else {
				  return String.valueOf(cell.getBooleanCellValue());
			  }}
			  
			catch (Exception e){
				e.printStackTrace();
				return "No matched value"; }
				
			}
			
			public  String getCellDataWithColNum(String sheetName,  int rowNum, int colNum) {
				try {
		
					
				  sheet=workbook.getSheet(sheetName);
				  row=sheet.getRow(rowNum);
				  cell = row.getCell(colNum);
				 
				  
				  if(cell.getCellTypeEnum()==CellType.STRING) {
					  return cell.getStringCellValue();
				  }
				  else if(cell.getCellTypeEnum()==CellType.NUMERIC || cell.getCellTypeEnum() == CellType.FORMULA) {
					  String cellValue=String.valueOf(cell.getNumericCellValue());
					  if(HSSFDateUtil.isCellDateFormatted(cell)) {
						  DateFormat dt= new SimpleDateFormat("dd/MM/yy");
						  Date date=cell.getDateCellValue();
						  cellValue=dt.format(date);
					  }
					  return cellValue;
				  }
				  else if(cell.getCellTypeEnum()==CellType.BLANK) {
					  return "";
				  }
				  else {
					  return String.valueOf(cell.getBooleanCellValue());
				  }}
				  
				catch (Exception e){
					e.printStackTrace();
					return "No matched value";
					
				}
			
		
}}