package com.cards.report;

import java.io.File;
import java.io.FileOutputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteExcelForCard {

	public static void main(String[] args) {

		// Blank workbook
		XSSFWorkbook workbook = new XSSFWorkbook();

		// Create a blank sheet
		XSSFSheet sheet = workbook.createSheet("Cards Data");

		// Create a new font and alter it.
		XSSFFont font = workbook.createFont();
		font.setBold(true);

		// Set font into style
		XSSFCellStyle style = workbook.createCellStyle();
		style.setFont(font);

		// This data needs to be written (Object[])
		Map<String, Object[]> data = new TreeMap<String, Object[]>();

		Date date = Calendar.getInstance().getTime();

		DateFormat dateFormat = new SimpleDateFormat("yyyy-mm-dd ");
		String strDate = dateFormat.format(date);

		Calendar c = Calendar.getInstance();

		c.setTime(date);
		c.add(Calendar.DAY_OF_MONTH, 7);
		String endDate = dateFormat.format(c.getTime());
		data.put("1", new Object[] { strDate + "-" + endDate, "Laborie", "Mon Repos", "Teachers", "Choiseul", "Dennery",
				"Fond St Jacques", "Hospitality Wkrs" });
	

		Map<String, Integer> dbMap1 = null;	
		dbMap1=getReultData(dbMap1,"jdbc:oracle:thin:@localhost:1521:xe","system", "oracle");

		


		Map<String, Integer> dbMap2 = new HashMap<String, Integer>();

		dbMap2.put("Balance inquiry", 10);
		dbMap2.put("Cards Active", 20);
		dbMap2.put("Cards Issued", 55);
		dbMap2.put("Cards withdrawal Rev", 40);
		dbMap2.put("Cash withdrawal", 10);
		dbMap2.put("Deposits", 50);
		dbMap2.put("Funds Transfer", 15);

		Map<String, Integer> dbMap3 = new HashMap<String, Integer>();

		dbMap3.put("Balance inquiry", 10);
		dbMap3.put("Cards Active", 20);
		dbMap3.put("Cards Issued", 55);
		dbMap3.put("Cards withdrawal Rev", 40);
		dbMap3.put("Cash withdrawal", 10);
		dbMap3.put("Deposits", 50);
		dbMap3.put("Funds Transfer", 15);

		Map<String, Integer> dbMap4 = new HashMap<String, Integer>();

		dbMap4.put("Balance inquiry", 10);
		dbMap4.put("Cards Active", 20);
		dbMap4.put("Cards Issued", 55);
		dbMap4.put("Cards withdrawal Rev", 40);
		dbMap4.put("Cash withdrawal", 10);
		dbMap4.put("Deposits", 50);
		dbMap4.put("Funds Transfer", 15);

		Map<String, Integer> dbMap5 = new HashMap<String, Integer>();

		dbMap5.put("Balance inquiry", 10);
		dbMap5.put("Cards Active", 20);
		dbMap5.put("Cards Issued", 55);
		dbMap5.put("Cards withdrawal Rev", 40);
		dbMap5.put("Cash withdrawal", 10);
		dbMap5.put("Deposits", 50);
		dbMap5.put("Funds Transfer", 15);

		Map<String, Integer> dbMap6 = new HashMap<String, Integer>();

		dbMap6.put("Balance inquiry", 14);
		dbMap6.put("Cards Active", 20);
		dbMap6.put("Cards Issued", 55);
		dbMap6.put("Cards withdrawal Rev", 40);
		dbMap6.put("Cash withdrawal", 10);
		dbMap6.put("Deposits", 50);
		dbMap6.put("Funds Transfer", 15);

		Map<String, Integer> dbMap7 = new HashMap<String, Integer>();

		dbMap7.put("Balance inquiry", 30);
		dbMap7.put("Cards Active", 20);
		dbMap7.put("Cards Issued", 9);
		dbMap7.put("Cards withdrawal Rev", 40);
		dbMap7.put("Cash withdrawal", 5);
		dbMap7.put("Deposits", 50);
		dbMap7.put("Funds Transfer", 4);

		data.put("2",
				new Object[] { "Balance inquiry", dbMap1.get("Balance inquiry"), dbMap2.get("Balance inquiry"),
						dbMap3.get("Balance inquiry"), dbMap4.get("Balance inquiry"), dbMap5.get("Balance inquiry"),
						dbMap6.get("Balance inquiry"), dbMap7.get("Balance inquiry") });
		data.put("3",
				new Object[] { "Cards Active", dbMap1.get("Cards Active"), dbMap2.get("Cards Active"),
						dbMap3.get("Cards Active"), dbMap4.get("Cards Active"), dbMap5.get("Cards Active"),
						dbMap6.get("Cards Active"), dbMap7.get("Cards Active") });
		data.put("4",
				new Object[] { "Cards Issued", dbMap1.get("Cards Issued"), dbMap2.get("Cards Issued"),
						dbMap3.get("Cards Issued"), dbMap4.get("Cards Issued"), dbMap5.get("Cards Issued"),
						dbMap6.get("Cards Issued"), dbMap7.get("Cards Issued") });
		data.put("5",
				new Object[] { "Cards withdrawal Rev", dbMap1.get("Cards withdrawal Rev"),
						dbMap2.get("Cards withdrawal Rev"), dbMap3.get("Cards withdrawal Rev"),
						dbMap4.get("Cards withdrawal Rev"), dbMap5.get("Cards withdrawal Rev"),
						dbMap6.get("Cards withdrawal Rev"), dbMap7.get("Cards withdrawal Rev") });
		data.put("6",
				new Object[] { "Cash withdrawal", dbMap1.get("Cash withdrawal"), dbMap2.get("Cash withdrawal"),
						dbMap3.get("Cash withdrawal"), dbMap4.get("Cash withdrawal"), dbMap5.get("Cash withdrawal"),
						dbMap6.get("Cash withdrawal"), dbMap7.get("Cash withdrawal") });
		data.put("7", new Object[] { "Deposits", dbMap1.get("Deposits"), dbMap2.get("Deposits"), dbMap3.get("Deposits"),
				dbMap4.get("Deposits"), dbMap5.get("Deposits"), dbMap6.get("Deposits"), dbMap7.get("Deposits") });
		data.put("8",
				new Object[] { "Funds Transfer", dbMap1.get("Funds Transfer"), dbMap2.get("Funds Transfer"),
						dbMap3.get("Funds Transfer"), dbMap4.get("Funds Transfer"), dbMap5.get("Funds Transfer"),
						dbMap6.get("Funds Transfer"), dbMap7.get("Funds Transfer") });

		// Iterate over data and write to sheet
		Set<String> keyset = data.keySet();
		int rownum = 0;
		for (String key : keyset) {
			Row row = sheet.createRow(rownum++);
			Object[] objArr = data.get(key);
			int cellnum = 0;
			for (Object obj : objArr) {
				Cell cell = row.createCell(cellnum++);
				if (obj instanceof String) {

					cell.setCellValue((String) obj);

				} else if (obj instanceof Integer) {
					cell.setCellValue((Integer) obj);

				}

			}
		}

		try {
			// Write the workbook in file system
			FileOutputStream out = new FileOutputStream(new File("cards_repo.xlsx"));
			workbook.write(out);
			out.close();

			System.out.println("card_repo.xlsx written successfully on disk.");

		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	private static Map<String, Integer> getReultData(Map<String, Integer> dbMap, String connecitonUrl, String userName,
			String password) {
		
		dbMap=new HashMap<String,Integer>();
		
		Connection con=null;
		ResultSet rs=null;
		Statement stmt =null;
		try {
			Class.forName("oracle.jdbc.driver.OracleDriver");

			 con = DriverManager.getConnection(connecitonUrl,userName, password);
			 stmt = con.createStatement();
			
			rs = stmt.executeQuery("select desc,count from emp");
			while (rs.next()) {

				dbMap.put(rs.getString("desc"), rs.getInt("count"));

			}

		} catch (ClassNotFoundException e1) { 
			e1.printStackTrace();
		} catch (SQLException e) {
			e.printStackTrace();
		}finally {
			try {
				con.close();
				rs.close();
				stmt.close();
			} catch (SQLException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			
		}
		return dbMap;
	}



}
