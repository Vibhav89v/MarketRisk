package test;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import common.AutomationConstants;

/**
 * A very simple program that writes some data to an Excel file
 * using the Apache POI library.
 * @author www.codejava.net
 *
 */
public class SetExcelData implements AutomationConstants{

	public static void main(String[] args) throws IOException 
	{

		
	}
	
	public static void writeExcelData(String sheetName, String rowName, String colName) throws FileNotFoundException, IOException
	{
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet(sheetName);
		Object[][] bookData = 
			{{rowName, colName},
					
			};

		int rowCount = 0;
		for (Object[] aBook : bookData) {
			Row row = sheet.createRow(++rowCount);
			int columnCount = 0;
			for (Object field : aBook) {
				Cell cell = row.createCell(++columnCount);
				if (field instanceof String) {
					cell.setCellValue((String) field);
				} else if (field instanceof Integer) {
					cell.setCellValue((Integer) field);
				}}}

		try (FileOutputStream outputStream = new FileOutputStream(OUTPUT_PATH)) {
			workbook.write(outputStream);
		}
	}
}
