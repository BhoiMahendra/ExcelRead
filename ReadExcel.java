package com;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;




public class ReadExcel {

	public static void main(String args[]) throws IOException {
		
		FileInputStream fis  = new FileInputStream(new File("C:\\abc.xls"));
		int count =0;
		int total = 0;
		
		FileWriter fw=new FileWriter("C:\\abc.txt");

		try {
			HSSFWorkbook wb=new HSSFWorkbook(fis);   
			
			
			HSSFSheet sheet=wb.getSheetAt(0);  
			 
			
			  for (Row row: sheet) {
				  
				  System.out.println("\t month \t"+ row.getCell(0).toString().substring(3, 6));
				  System.out.println("\t year \t"+ row.getCell(0).toString().substring(7, 9));
				  String month =  row.getCell(0).toString().substring(3, 6);
				  String year = row.getCell(0).toString().substring(7, 9).toString();
				
				  if(month.equalsIgnoreCase("JUL") && year.equalsIgnoreCase("20")) {
					  count ++;
					  total++;
					  fw.write("Delete from ftdmasterkeyvalue where processid ="+ "'W901'"+" AND TXREFNO = "+"'"+row.getCell(1).toString()+"'" +" and value is null; \n");
				  }
				  if(count == 500) {
					  fw.write("COMMIT;\n");
					  
					  count=0;
				  }
				  
			  
			  }
			  fw.write("COMMIT;\n");
			  fw.close();
			 
			
		} catch (IOException e) {
			e.printStackTrace();
		}
		
		
		 
		 
	}
	
}
