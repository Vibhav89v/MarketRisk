package test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.management.ManagementFactory;
import java.lang.management.MemoryMXBean;
import java.lang.management.MemoryUsage;
import java.util.Iterator;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.Assert;
import org.testng.asserts.SoftAssert;

import common.AutomationConstants;

public class CompareExcelFiles implements AutomationConstants {

	private static final int MEGABYTE = 0;
	private static final int GIGABYTE = 0;
	public static String colName;
	public static String rowName;
	public static String sourceSheet = "Temp_FixedInstrument_Results";
	public static String targetSheet = "ST_M_FixedInstrument_Results";
	public static int rowsInSourceSheet;
	public static int columnsInSourceSheet;
	public static int rowsInTargetSheet;
	public static int columnsInTargetSheet;
	public static int sheetCounts;
	public static String v1;
	public static String v2;


	public static void verifySheetsInExcelFilesHaveSameRowsAndColumns(Workbook workbook1) {

		//		public void verifySheetsInExcelFilesHaveSameRowsAndColumns(Workbook workbook1, Workbook workbook2) {

		System.out.println("Verifying if both sheets have same number of rows and columns.............");

		sheetCounts = workbook1.getNumberOfSheets();

		for (int i = 0; i < sheetCounts; i++) 
		{
			Sheet s1 = workbook1.getSheet(sourceSheet);
			Sheet s2 = workbook1.getSheet(targetSheet);
			rowsInSourceSheet = s1.getPhysicalNumberOfRows();
			rowsInTargetSheet = s2.getPhysicalNumberOfRows();

			Assert.assertEquals(rowsInSourceSheet, rowsInTargetSheet, "Sheets have different count of rows.." );

			Iterator<Row> rowInSheet1 = s1.rowIterator();
			Iterator<Row> rowInSheet2 = s2.rowIterator();
			while (rowInSheet1.hasNext()) {
				int columnsInSourceSheet1 = rowInSheet1.next().getPhysicalNumberOfCells();
				int columnsInSourceSheet2 = rowInSheet2.next().getPhysicalNumberOfCells();
				Assert.assertEquals(columnsInSourceSheet1, columnsInSourceSheet2, "Sheets have different count of columns..");
			}
		}
	}

	public static void verifyBlankCellsInSheet(Workbook workbook1) throws EncryptedDocumentException, InvalidFormatException, IOException {


		System.out.println();

		System.out.println(" *********************  VERIFIES IF THERE IS NO BLANK CELL IN THE EXCEL CELL **************************************");

		Sheet s1 = workbook1.getSheet(sourceSheet);
		Sheet s2 = workbook1.getSheet(targetSheet);


		rowsInSourceSheet = s1.getPhysicalNumberOfRows();

		SoftAssert softAssertion= new SoftAssert();

		for (int j = 0; j < rowsInSourceSheet; j++) 
		{
			// Iterating through each cell
			int columnsInSourceSheet = s1.getRow(j).getPhysicalNumberOfCells();

			for (int k = 1; k <= columnsInSourceSheet -1; k++)
			{
				// Getting individual cell
				Cell c1 = s1.getRow(j).getCell(k);
				Cell c2 = s2.getRow(j).getCell(k);
				CellType ct1 = c1.getCellTypeEnum();
				CellType ct2 = c2.getCellTypeEnum();


				if ( !ct1.equals(ct2))
				{					
					colName = s2.getRow(0).getCell(k).getStringCellValue();
					rowName = s2.getRow(j).getCell(1).getStringCellValue();

					System.out.println("CELL CORDINATE IS  : " + "row - " + j + ", " + rowName + "; column- " + k + ", " + colName);

					System.out.println("CELL TYPE IN " + s1.getSheetName().toUpperCase() + " IS : " + ct1);
					System.out.println("CELL TYPE IN " + s2.getSheetName().toUpperCase() + " IS : " + ct2 + " ; change the cell type to " + ct1);

					System.out.println();

					softAssertion.assertAll();

				}
			}
		}

	}

	@SuppressWarnings("deprecation")
	public static void verifyDataInExcelBookAllSheets(Workbook workbook1) throws EncryptedDocumentException, InvalidFormatException, IOException 
	{

		System.out.println("Verifying if both work books have same data.............");

		MemoryMXBean memoryBean = ManagementFactory.getMemoryMXBean();
		byte[] bytes = new byte[GIGABYTE*500];
		//	    for (int i=1; i <= 100; i++) {
		try {
			int rowCount = 0;
			int columnCount = 0;
			// Since we have already verified that both work books have same number of sheets so iteration can be done against any workbook sheet count
			int sheetCounts = workbook1.getNumberOfSheets();

			// Get sheet at same index of both work books

			Sheet s1 = workbook1.getSheet(sourceSheet);
			Sheet s2 = workbook1.getSheet(targetSheet);

			XSSFWorkbook workbook = new XSSFWorkbook();
			XSSFSheet sheet = workbook.createSheet("Output Report");
			System.out.println("*********** Sheet Name : "+s1.getSheetName()+"*************");
			// Iterating through each row

			rowsInSourceSheet = s1.getPhysicalNumberOfRows();

			int createRowCount = 1;
			int serialNum = 1;

			SoftAssert softAssertion= new SoftAssert();


			for (int j = 0; j < rowsInSourceSheet; j++) 
			{
				// Iterating through each cell
				columnsInSourceSheet = s1.getRow(j).getPhysicalNumberOfCells();

				for (int k = 1; k <= columnsInSourceSheet -1; k++)
				{
					Cell cell1 = s1.getRow(j).getCell(k);
					Cell cell2 = s2.getRow(j).getCell(k);
					String column_letter = CellReference.convertNumToColString(cell2.getColumnIndex());
					String c1 = "";
					String c2 = "";
					String orderNum = new DataFormatter().formatCellValue(s2.getRow(j).getCell(0));

					if(cell1.getCellTypeEnum() == CellType.NUMERIC) 
					{
						if (DateUtil.isCellDateFormatted(cell1) | DateUtil.isCellDateFormatted(cell2)) 
						{
							//							System.out.println("DATE CELL VALUE FOR SHEET 1: " + cell1.getDateCellValue());
							//							System.out.println("DATE CELL VALUE FOR SHEET 2: " + cell2.getDateCellValue());
							// Need to use DataFormatter to get data in given style otherwise it will come as time stamp
							DataFormatter df = new DataFormatter();
							c1 = df.formatCellValue(cell1);
							c2 = df.formatCellValue(cell2);

							//							System.out.println("Date Format Cell Value Sheet 1 : "+c1);
							//							System.out.println("Date Format Cell Value Sheet 2 : "+c2);
							if (!c1.equals(c2))
							{
								colName = s2.getRow(0).getCell(k).getStringCellValue();
								rowName = s2.getRow(j).getCell(2).getStringCellValue();

								System.out.println(orderNum + column_letter + " DATES DID NOT MATCH FOR " + rowName + " and " + colName + " , difference is "+ c1 + " === "+ c2);

								Sheet s3 = workbook1.getSheet("Output Report"); 

								Row row3 = s3.createRow(createRowCount++);
								row3.createCell(0).setCellValue(serialNum);
								row3.createCell(1).setCellValue(orderNum);
								row3.createCell(2).setCellValue(column_letter);
								row3.createCell(3).setCellValue(rowName);
								row3.createCell(4).setCellValue(colName);
								row3.createCell(5).setCellValue(c1);
								row3.createCell(6).setCellValue(c2);
								serialNum++;

							}
						}

						else 
						{
							c1 = NumberToTextConverter.toText(cell1.getNumericCellValue());
							c2 = NumberToTextConverter.toText(cell2.getNumericCellValue());
							if ( !c1.equals(c2))
							{
								colName = s2.getRow(0).getCell(k).getStringCellValue();
								cell2 = s2.getRow(j).getCell(2);
								// Getting individual cell
								rowName = s2.getRow(j).getCell(2).toString();
								if(cell2.getCellTypeEnum() == CellType.NUMERIC) 
								{
									rowName = NumberToTextConverter.toText(cell2.getNumericCellValue());
								}
								System.out.println(orderNum + column_letter + " NUMERIC VALUE DID NOT MATCH FOR " + rowName + " and " + colName +" , difference is "+ c1 + " === "+ c2);

								Sheet s3 = workbook1.getSheet("Output Report"); 

								Row row3 = s3.createRow(createRowCount++);
								row3.createCell(0).setCellValue(serialNum);
								row3.createCell(1).setCellValue(orderNum);
								row3.createCell(2).setCellValue(column_letter);
								row3.createCell(3).setCellValue(rowName);
								row3.createCell(4).setCellValue(colName);
								row3.createCell(5).setCellValue(c1);
								row3.createCell(6).setCellValue(c2);

								serialNum++;
							}
						}
					}

					else 
					{
						c1 = cell1.getStringCellValue();
						c2 = cell2.getStringCellValue();

						if (!c1.equals(c2))
						{
							colName = s2.getRow(0).getCell(k).getStringCellValue();
							cell2 = s2.getRow(j).getCell(2);
							// Getting individual cell
							rowName = s2.getRow(j).getCell(2).toString();
							if(cell2.getCellTypeEnum() == CellType.NUMERIC) 
							{
								rowName = NumberToTextConverter.toText(cell2.getNumericCellValue());
							}
							System.out.println(orderNum + column_letter + " STRING VALUES DID NOT MATCH FOR " + rowName + " and " + colName + " , difference is "+ c1 + " === "+ c2);

							Sheet s3 = workbook1.getSheet("Output Report"); 

							Row row3 = s3.createRow(createRowCount++);
							row3.createCell(0).setCellValue(serialNum);
							row3.createCell(1).setCellValue(orderNum);
							row3.createCell(2).setCellValue(column_letter);
							row3.createCell(3).setCellValue(rowName);
							row3.createCell(4).setCellValue(colName);
							row3.createCell(5).setCellValue(c1);
							row3.createCell(6).setCellValue(c2);

							serialNum++;
						}
					}
				}
			}

			FileOutputStream outputStream = new FileOutputStream(INPUT_PATH);
			workbook1.write(outputStream);
			workbook1.close();
			System.out.println("Hurray! Both sheets have compared successfully.");
		}
		catch (Exception e) 
		{
			e.printStackTrace();
		} 
		catch (OutOfMemoryError e) 
		{
			MemoryUsage heapUsage = memoryBean.getHeapMemoryUsage();
			long maxMemory = heapUsage.getMax() / MEGABYTE;
			long usedMemory = heapUsage.getUsed() / MEGABYTE;
			System.out.println( " : Memory Use :" + usedMemory + "M/" +maxMemory+"M");
		}
	}

	public static void main (String[] args) throws EncryptedDocumentException, InvalidFormatException, FileNotFoundException, IOException
	{
		CompareExcelFiles.verifyDataInExcelBookAllSheets(WorkbookFactory.create(new FileInputStream(INPUT_PATH)));
		//		CompareExcelFiles.verifySheetsInExcelFilesHaveSameRowsAndColumns( WorkbookFactory.create(new FileInputStream(INPUT_PATH)));
		//	CompareExcelFiles.verifyBlankCellsInSheet( WorkbookFactory.create(new FileInputStream(INPUT_PATH)));
	}
}