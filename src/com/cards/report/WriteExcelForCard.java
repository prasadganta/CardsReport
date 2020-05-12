package com.cards.report;


import java.io.File;
import java.io.FileOutputStream;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteExcelForCard {


		public static void main(String[] args) 
		{
			//Blank workbook
			XSSFWorkbook workbook = new XSSFWorkbook(); 
			
			//Create a blank sheet
			XSSFSheet sheet = workbook.createSheet("Cards Data");
			 
			//This data needs to be written (Object[])
			Map<String, Object[]> data = new TreeMap<String, Object[]>();
			
			Date date = Calendar.getInstance().getTime();
			
			DateFormat dateFormat = new SimpleDateFormat("yyyy-mm-dd ");
			String strDate = dateFormat.format(date);
			
			Calendar c = Calendar.getInstance();
			
			c.setTime(date);
			c.add(Calendar.DAY_OF_MONTH, 7);  
			String endDate = dateFormat.format(c.getTime()); 
			data.put("1", new Object[] {strDate +"-"+endDate,"Laborie", "Mon Repos", "Teachers","Choiseul","Dennery","Fond St Jacques", "Hospitality Wkrs"});
			
			 // need to connect DB and loop it 
			
			for(int i=0;i<5;i++) {
				
				
			data.put(String.valueOf(i), new Object[] {"Cards Issued",1, 22, 72,11,10,0,0});
			}
		
			 
			//Iterate over data and write to sheet
			Set<String> keyset = data.keySet();
			int rownum = 0;
			for (String key : keyset)
			{
			    Row row = sheet.createRow(rownum++);
			    Object [] objArr = data.get(key);
			    int cellnum = 0;
			    for (Object obj : objArr)
			    {
			       Cell cell = row.createCell(cellnum++);
			       if(obj instanceof String)
			            cell.setCellValue((String)obj);
			        else if(obj instanceof Integer)
			            cell.setCellValue((Integer)obj);
			    }
			}
			try 
			{
				//Write the workbook in file system
			    FileOutputStream out = new FileOutputStream(new File("cards_repo.xlsx"));
			    workbook.write(out);
			    out.close();
			    
			    System.out.println("card_repo.xlsx written successfully on disk.");
			     
			} 
			catch (Exception e) 
			{
			    e.printStackTrace();
			}
		}

}
