package com.howtodoinjava.demo.poi;

import java.io.File;
import java.io.FileOutputStream;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;



public class WriteExcelDemo 
{
	public static void main(String[] args) 
	{
		//Blank workbook
		XSSFWorkbook workbook = new XSSFWorkbook(); 
		
		//Create a blank sheet
		XSSFSheet sheet = workbook.createSheet("Employee Data");
		 
		//This data needs to be written (Object[])
		Map<String, Person> data = new TreeMap<String, Person>();
		
		data.put("2", new Person(1, "Amit",123));
		data.put("3", new Person(2, "Lokesh",23));
		data.put("4", new Person(3, "John", 87));
		data.put("5", new Person(4, "Brian", 45));
		 
		//Iterate over data and write to sheet
		Set<String> keyset = data.keySet();
		int rownum   =  0;
		int cellnum  =  0;
		Iterator itr=keyset.iterator();
		while(itr.hasNext()) {
			
			String it=(String) itr.next();
			//System.out.println(it);
			
			Row row = sheet.createRow(rownum++);
			row.createCell(0).setCellValue(it);
			
			Person     p1    =     data.get(it);
			System.out.println(it +"     888       "+p1);
			Cell cell    =     row.createCell(cellnum++);
			row.createCell(0).setCellValue(  p1.getId());
			row.createCell(1).setCellValue(p1.getName());
			row.createCell(2).setCellValue( p1.getAge());
		}
		
		try 
		{
			//Write the workbook in file system
		    FileOutputStream out = new FileOutputStream(new File("howtodoinjava_demo.xlsx"));
		    workbook.write(out);
		    out.close();
		    
		    System.out.println("howtodoinjava_demo.xlsx written successfully on disk.");
		     
		} 
		catch (Exception e) 
		{
		    e.printStackTrace();
		}
	}
}
