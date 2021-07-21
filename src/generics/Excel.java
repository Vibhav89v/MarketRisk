package generics;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import common.AutomationConstants;

 

public class Excel implements AutomationConstants
{
	
	
//    public static void main(String[] args) throws EncryptedDocumentException, InvalidFormatException, IOException
//    {
//        FileInputStream fis=new FileInputStream(INPUT_PATH);
//        Workbook wb= WorkbookFactory.create(fis);
//        Sheet sh=wb.getSheet("trial");
//        
//        int lastRowNum = sh.getLastRowNum(); 
//        for(int i=0;i<=lastRowNum;i++)
//        {
//            Row row = sh.getRow(lastRowNum);
//            for(int j=0; j<=row.getLastCellNum();j++)
//            {
//            Cell cel = row.createCell(j);
//            cel.setCellValue("fuygr");
//            }
//        }
//        //sh.shiftRows(r, lastRowNum, 1);
//        //sh.createRow(2);    
//        
//        Row row = sh.getRow(lastRowNum);
//        Cell cel = row.createCell(0);
//        cel.setCellValue("fuygr");
//
//    
//    FileOutputStream fos = new FileOutputStream(INPUT_PATH);
//    wb.write(fos);
//    wb.close();
//        
//    }
    
    public static int getRowCount(String path,String sheet)
    {
        int r=0;
        try
        {
            r=WorkbookFactory.create(new FileInputStream(path)).getSheet(sheet).getLastRowNum();
        }
        catch(Exception e)
        {
            e.printStackTrace();
        }
        return r;
    }
    
//    public static String getSheet(String path)
//    {
//        String r=0;
//        try
//        {
//            r=WorkbookFactory.create(new FileInputStream(path)).getSheetAt(0);
//        }
//        catch(Exception e)
//        {
//            e.printStackTrace();
//        }
//        return r;
//    }
    
//    public static int getColumnCount(String path, String sheet)
//    {
//    	int c = 0;
//    	try
//    	{
//    		c=WorkbookFactory.create(new FileInputStream(path)).getSheet(sheet).getLastRowNum().
//    	}
//    }
    
    public static String getCellValue(String path,String sheet,int r,int c)
    {
        String v="";
        try
        {    
            CellType cs=WorkbookFactory.create(new FileInputStream(path)).getSheet(sheet).getRow(r).getCell(c).getCellTypeEnum();
            v=WorkbookFactory.create(new FileInputStream(path)).getSheet(sheet).getRow(r).getCell(c).toString();
            if(cs == CellType.NUMERIC)
            {
                //System.out.println("V:::"+ v);
                v = v.split("\\.")[0];
                //System.out.println("V:::"+ v);
            }
        }
        catch(Exception e)
        {
            e.printStackTrace();
        }
        return v;    
    }
    
    public static void setExcelData(String path, String sheet, int rowNum, int colNum,String data) 
            throws EncryptedDocumentException, InvalidFormatException, IOException
    {
        FileInputStream fis = new FileInputStream(path);
        Workbook wb = WorkbookFactory.create(fis);
        Sheet sh = wb.getSheet(sheet);
        
//        System.out.println(rowNum);
        	sh.createRow(rowNum);
//        	System.out.println("Row Created");
            Row row = sh.getRow(rowNum);
//           System.out.println(row);
           
            Cell cel = row.createCell(colNum);
            cel.setCellValue(data);

 

        
        FileOutputStream fos = new FileOutputStream(path);
        wb.write(fos);
        wb.close();
    }
    
    public static String shiftingRowsUp(String path,String sheet,int r) 
    {
    
        try
        {    
            FileInputStream fis=new FileInputStream(path);
            Workbook wb= WorkbookFactory.create(fis);
            Sheet sh=wb.getSheet(sheet);
            int lastRowNum = sh.getLastRowNum();
            if (r >= 0 && r < lastRowNum) {
                sh.shiftRows(r + 1, lastRowNum, -1);
            }
            if (r == lastRowNum) {
                Row removingRow = sh.getRow(r);
                if (removingRow != null) {
                    sh.removeRow(removingRow);
                }
            }
            
        
            FileOutputStream fos = new FileOutputStream(path);
            wb.write(fos);
            fos.flush();
            fos.close();
          
        }
        catch(Exception e)
        {
            e.printStackTrace();
        }
        return sheet;
        
        
    }
    
    public static String shiftingRowsDown(String path,String sheet,int r)
    {

 

        try
        {    
            FileInputStream fis=new FileInputStream(path);
            Workbook wb= WorkbookFactory.create(fis);
            Sheet sh=wb.getSheet(sheet);
            int lastRowNum = sh.getLastRowNum();
            sh.shiftRows(r, lastRowNum, 1);
            sh.createRow(2);    
            FileOutputStream fos = new FileOutputStream(path);
            wb.write(fos);
            fos.flush();
            fos.close();
            
          
        }
        catch(Exception e)
        {
            e.printStackTrace();
        }
        return sheet;
    
    }
    
    }
