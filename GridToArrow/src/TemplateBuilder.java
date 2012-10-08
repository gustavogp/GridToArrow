import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Calendar;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;


public class TemplateBuilder {
	//fields
	static Calendar todayNow = Calendar.getInstance();
	static Workbook wb = new HSSFWorkbook();
	static Workbook wbMix = new HSSFWorkbook();
	static Sheet sheet1 = wb.createSheet("Forecast");
	static Sheet sheetMix = wbMix.createSheet("Mix");
	
	public static void createTemplate() {
		FileOutputStream fileOut;
		CellStyle style1 = wb.createCellStyle();
		CellStyle style2 = wb.createCellStyle();
		CellStyle style3 = wb.createCellStyle();
		CellStyle style4 = wb.createCellStyle();
		
		//set styles, but not using yet
		style1.setFillForegroundColor(IndexedColors.LIGHT_ORANGE.getIndex());
	    style1.setFillPattern(CellStyle.SOLID_FOREGROUND);
	    style2.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
	    style2.setFillPattern(CellStyle.SOLID_FOREGROUND);
	    style3.setFillForegroundColor(IndexedColors.CORAL.getIndex());
	    style3.setFillPattern(CellStyle.SOLID_FOREGROUND);
	    style4.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
	    style4.setFillPattern(CellStyle.SOLID_FOREGROUND);
		
		//create header
		Row row = sheet1.createRow(0);
		row.createCell(0).setCellValue("Customer Name");
		row.createCell(1).setCellValue("Apple Part Nr");
		row.createCell(2).setCellValue("Date Updated");
		row.createCell(3).setCellValue("Region Code");
		row.createCell(4).setCellValue("Instock Percentage");
		row.createCell(5).setCellValue("Store Count");
		row.createCell(6).setCellValue("DC On Hand");
		row.createCell(7).setCellValue("Target WOS");
		row.createCell(8).setCellValue("Current Week Trend");
		row.createCell(9).setCellValue("Forecast Quarter");
		row.createCell(10).setCellValue("Week 1 Forecast");
		row.createCell(11).setCellValue("Week 2 Forecast");
		row.createCell(12).setCellValue("Week 3 Forecast");
		row.createCell(13).setCellValue("Week 4 Forecast");
		row.createCell(14).setCellValue("Week 5 Forecast");
		row.createCell(15).setCellValue("Week 6 Forecast");
		row.createCell(16).setCellValue("Week 7 Forecast");
		row.createCell(17).setCellValue("Week 8 Forecast");
		row.createCell(18).setCellValue("Week 9 Forecast");
		row.createCell(19).setCellValue("Week 10 Forecast");
		row.createCell(20).setCellValue("Week 11 Forecast");
		row.createCell(21).setCellValue("Week 12 Forecast");
		row.createCell(22).setCellValue("Week 13 Forecast");
		row.createCell(23).setCellValue("Week 14 Forecast");
		row.createCell(24).setCellValue("Week 1 JST");
		row.createCell(25).setCellValue("Week 2 JST");
		row.createCell(26).setCellValue("Week 3 JST");
		row.createCell(27).setCellValue("Week 4 JST");
		row.createCell(28).setCellValue("Week 5 JST");
		row.createCell(29).setCellValue("Week 6 JST");
		row.createCell(30).setCellValue("Week 7 JST");
		row.createCell(31).setCellValue("Week 8 JST");
		row.createCell(32).setCellValue("Week 9 JST");
		row.createCell(33).setCellValue("Week 10 JST");
		row.createCell(34).setCellValue("Week 11 JST");
		row.createCell(35).setCellValue("Week 12 JST");
		row.createCell(35).setCellValue("Week 13 JST");
		row.createCell(37).setCellValue("Week 14 JST");
		row.createCell(38).setCellValue("Week 1 Req");
		row.createCell(39).setCellValue("Week 2 Req");
		row.createCell(40).setCellValue("Week 3 Req");
		row.createCell(41).setCellValue("Week 4 Req");
		row.createCell(42).setCellValue("Week 5 Req");
		row.createCell(43).setCellValue("Week 6 Req");
		row.createCell(44).setCellValue("Week 7 Req");
		row.createCell(45).setCellValue("Week 8 Req");
		row.createCell(46).setCellValue("Week 9 Req");
		row.createCell(47).setCellValue("Week 10 Req");
		row.createCell(48).setCellValue("Week 11 Req");
		row.createCell(49).setCellValue("Week 12 Req");
		row.createCell(50).setCellValue("Week 13 Req");
		row.createCell(51).setCellValue("Week 14 Req");
		
		//create and save the xls file
		try {
			fileOut = new FileOutputStream(GTAGUI.inputPath.getParent() + "/ForecastTemplate_" + String.valueOf((todayNow.get(Calendar.MONTH)+1)) + String.valueOf(todayNow.get(Calendar.DAY_OF_MONTH)) + 
					String.valueOf(todayNow.get(Calendar.HOUR_OF_DAY)) + String.valueOf(todayNow.get(Calendar.MINUTE)) + ".xls");
			wb.write(fileOut);
			fileOut.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
			GTAGUI.generalMessage("Error saving Template file: " + e.getMessage());
		} catch (IOException e) {
			e.printStackTrace();
			GTAGUI.generalMessage("Error saving Template file" + e.getMessage());
		} 
		
	}
	
	public static void createMixTemplate(Map<Integer, HashMap<String, Double>> mixTemp) {
		FileOutputStream fileOut;
		int rCount = 0;
		Row row;
		
		//create header
				row = sheetMix.createRow(rCount);
				rCount++;
				row.createCell(0).setCellValue("Customer Name");
				row.createCell(1).setCellValue("Apple Part Nr");
				row.createCell(2).setCellValue("Date Updated");
				row.createCell(3).setCellValue("Region Code");
				row.createCell(4).setCellValue("Instock Percentage");
				row.createCell(5).setCellValue("Store Count");
				row.createCell(6).setCellValue("DC On Hand");
				row.createCell(7).setCellValue("Target WOS");
				row.createCell(8).setCellValue("Current Week Trend");
				row.createCell(9).setCellValue("Forecast Quarter");
				row.createCell(10).setCellValue("Week 1 Forecast");
				row.createCell(11).setCellValue("Week 2 Forecast");
				row.createCell(12).setCellValue("Week 3 Forecast");
				row.createCell(13).setCellValue("Week 4 Forecast");
				row.createCell(14).setCellValue("Week 5 Forecast");
				row.createCell(15).setCellValue("Week 6 Forecast");
				row.createCell(16).setCellValue("Week 7 Forecast");
				row.createCell(17).setCellValue("Week 8 Forecast");
				row.createCell(18).setCellValue("Week 9 Forecast");
				row.createCell(19).setCellValue("Week 10 Forecast");
				row.createCell(20).setCellValue("Week 11 Forecast");
				row.createCell(21).setCellValue("Week 12 Forecast");
				row.createCell(22).setCellValue("Week 13 Forecast");
				row.createCell(23).setCellValue("Week 14 Forecast");
				
				//create SKU column
				for(String k : mixTemp.get(1).keySet()) {
					row = sheetMix.createRow(rCount);
					rCount++;
					row.createCell(1).setCellValue(k);
				}
				//add mix data
				for (int wk = 1; wk < mixTemp.size() + 1; wk++) {
					try{
						for(Row r : sheetMix) {
							if (r.getRowNum() > 0) {
								r.createCell(9 + wk).setCellValue(mixTemp.get(wk).get(r.getCell(1).getStringCellValue()));
							}
							
						}
					} catch (NullPointerException e) {
						//e.printStackTrace();
					}
					
				}
				
				//create and save the xls file
				try {
					fileOut = new FileOutputStream(GTAGUI.inputPath.getParent() + "/ForecastMix_" + String.valueOf((todayNow.get(Calendar.MONTH)+1)) + String.valueOf(todayNow.get(Calendar.DAY_OF_MONTH)) + 
							String.valueOf(todayNow.get(Calendar.HOUR_OF_DAY)) + String.valueOf(todayNow.get(Calendar.MINUTE)) + ".xls");
					wbMix.write(fileOut);
					fileOut.close();
				} catch (FileNotFoundException e) {
					e.printStackTrace();
					GTAGUI.generalMessage("Error saving Template file: " + e.getMessage());
				} catch (IOException e) {
					e.printStackTrace();
					GTAGUI.generalMessage("Error saving Template file" + e.getMessage());
				}
				
				
	}
	
	
}
