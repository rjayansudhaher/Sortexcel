package com.excel.reading;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Sort {

	public static void main(String[] args) throws IOException {
		
		// Need to change the path accordingly
		File file = new File("C:\\Users\\jayar29\\Desktop\\jhry15.xlsx");
		   FileInputStream fIP = new FileInputStream(file);
		   XSSFWorkbook workbook = new XSSFWorkbook(fIP);
		   
		   workbook=sortCells(workbook);
		   
		   FileOutputStream out = new FileOutputStream( 
				      new File("C:\\Users\\jayar29\\Desktop\\JHRY_15_test.xlsx"));
				      workbook.write(out);
				      out.close();

	}
	
	private static XSSFWorkbook sortCells(XSSFWorkbook workbook) {
		
		 int sheetCount = workbook.getNumberOfSheets();
		 for(int i=sheetCount-1;i>=0 ;i--){
			 
			 int sortingcol=271;
	        
	        XSSFSheet sheet = workbook.getSheetAt(i);
	        Row row1=null;
	        Row row2=null;
	        
	        boolean sorting = true;
	        int lastRow = sheet.getLastRowNum();
	        
	        
	        try{
	        	for(int j=1; j<=sheet.getLastRowNum();j++){
	        
	        	row1=sheet.getRow(j);
	        	Cell ce1=row1.getCell(sortingcol);
	        	
	        	for (int k=j+1;k<=sheet.getLastRowNum();k++)
	        	{
	        		row2=sheet.getRow(k);
	        		Cell ce2=row2.getCell(sortingcol);
	        		
	        		System.out.println("@@"+ce1.toString());
	                System.out.println("##"+ce2.toString());
	        		if(ce2.toString().compareToIgnoreCase(ce1.toString())<0)
	        		{
	        			System.out.println("Soort");
	        			for(int l=0;l<row1.getLastCellNum();l++)
	        			{
	        				if(row1.getCell(l) !=null && row2.getCell(l) !=null){
	        				String s=row1.getCell(l).toString();
	        				row1.getCell(l).setCellValue(row2.getCell(l).toString());
	        				row2.getCell(l).setCellValue(s);
	        				}
	        				
	        			}
	        			
	        			
	        		}
	        	}
	        	
	        	
			}
	        }catch(Exception e)
	        {
	        	//e.printStackTrace();
	        	continue;
	        }
	        
	        
		 }
		
		return workbook;
	}

}
