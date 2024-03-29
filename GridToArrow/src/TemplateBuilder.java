import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Calendar;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.formula.FormulaParseException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;


public class TemplateBuilder {
	//fields
	static Calendar todayNow = Calendar.getInstance();
	static Workbook wb = new HSSFWorkbook();
	static Sheet sheetForByAcc = wb.createSheet("Forecast By Account");
	static Sheet sheetMix = wb.createSheet("Mix");
	static Sheet sheet1 = wb.createSheet("Judged Forecast");
	static FileOutputStream fileOut;
	static int rCountFBA = 0;
	static int previousLastRow = 0;
	static int oldPreviousLastRow = 0;
	static int rCountMix = 0;
	static int previousLastRowMix = 1;
	static int rCountFore = 0;
	static int previousLastRowFore = 0;
	static int rowsToLeap = 2;
	static Map<Integer,Integer> RTLMap = new HashMap<Integer,Integer>();
	static int beforeThisRow = 0;
	static Map<Integer,Integer> BTRMap = new HashMap<Integer,Integer>();
	
	/**
	 * Forecast by account BEFORE CSAM judgment, should be the first tab in the file
	 * @param weeklyMapsSF
	 * @param isFirst
	 * @param isLast
	 */
	public static void FrcstByAccntBySubF(String name, Map<Integer, LinkedHashMap<String, Integer>> weeklyMapsSF, boolean isFirst, boolean isLast) {
		Row row;
		
		//create header only if isFirst
		if (isFirst) {
			
			row = sheetForByAcc.createRow(rCountFBA);
			rCountFBA++;
			row.createCell(0).setCellValue("Customer Name");
			row.createCell(1).setCellValue("Sub LOB");
			row.createCell(2).setCellValue("Week 1 ");
			row.createCell(3).setCellValue("Week 2 ");
			row.createCell(4).setCellValue("Week 3 ");
			row.createCell(5).setCellValue("Week 4 ");
			row.createCell(6).setCellValue("Week 5 ");
			row.createCell(7).setCellValue("Week 6 ");
			row.createCell(8).setCellValue("Week 7 ");
			row.createCell(9).setCellValue("Week 8 ");
			row.createCell(10).setCellValue("Week 9 ");
			row.createCell(11).setCellValue("Week 10 ");
			row.createCell(12).setCellValue("Week 11 ");
			row.createCell(13).setCellValue("Week 12 ");
			row.createCell(14).setCellValue("Week 13 ");
			row.createCell(15).setCellValue("Week 14 ");
		}	
		//create SubFamily column
		for(String k : weeklyMapsSF.get(1).keySet()) { //could pick any of the weeks, using wk 1 here
			row = sheetForByAcc.createRow(rCountFBA);
			row.createCell(0).setCellValue(name);
			row.createCell(1).setCellValue(k);
			rCountFBA++;
		}
		//add forecast data
		for (int wk = 1; wk < weeklyMapsSF.size() + 1; wk++) {
			try{
				for(Row r : sheetForByAcc) {
					if (r.getRowNum() > previousLastRow) {
						r.createCell(1 + wk).setCellValue(weeklyMapsSF.get(wk).get(r.getCell(1).getStringCellValue()));
					}
					
				}
			} catch (NullPointerException e) {
				//e.printStackTrace();
			}
			
		}
		
		for( Cell c : sheetForByAcc.getRow(2)) {
			sheetForByAcc.autoSizeColumn(c.getColumnIndex());
		}
		
		sheetForByAcc.createFreezePane(2, 1);
		
		//update the previousLastRow, subtract 1 since we had added 1 and didn't "use" the row yet
		oldPreviousLastRow = previousLastRow;
		previousLastRow = rCountFBA - 1;
			
	}
	
	/**
	 * Mix sheet, should be the second tab in the file. Should include column for CSA judgment
	 * @param mixTemp
	 */
	public static void createMixTemplate(String name, Map<Integer, LinkedHashMap<String, Double>> weeklyMapsSku, boolean isFirst, boolean isLast) {
		Row row;
		DataFormat df;
		CellStyle percentageStyle, percentageStyle2, centerStyle, percentBoldStyle;
		
		//create format styles and font
		df = wb.createDataFormat();
		percentageStyle = wb.createCellStyle();
		percentageStyle.setDataFormat(df.getFormat("0.00%"));
		
		percentageStyle2 = wb.createCellStyle();
		percentageStyle2.setDataFormat(df.getFormat("0.00%"));
		percentageStyle2.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
		percentageStyle2.setFillPattern(CellStyle.SOLID_FOREGROUND);
		percentageStyle2.setBorderRight(CellStyle.BORDER_THICK);
		
		centerStyle = wb.createCellStyle();
		centerStyle.setAlignment(CellStyle.ALIGN_CENTER);
		
		Font f = wb.createFont();
		f.setBoldweight(Font.BOLDWEIGHT_BOLD);
		percentBoldStyle = wb.createCellStyle();
		percentBoldStyle.setFont(f);
		percentBoldStyle.setDataFormat(df.getFormat("0.00%"));
		percentBoldStyle.setBorderRight(CellStyle.BORDER_THICK);
		
		//create header, merge header cells, add row bellow header (subheader)
		if(isFirst) {
			row = sheetMix.createRow(rCountMix);
			rCountMix++;
			row.createCell(0).setCellValue("Customer Name");
			row.createCell(1).setCellValue("Apple Part Nr");
			row.createCell(2).setCellValue("Week 1");
			sheetMix.addMergedRegion(new CellRangeAddress(rCountMix-1, rCountMix-1, 2, 4));
			row.getCell(2).setCellStyle(centerStyle);
			row.createCell(5).setCellValue("Week 2");
			sheetMix.addMergedRegion(new CellRangeAddress(rCountMix-1, rCountMix-1, 5, 7));
			row.getCell(5).setCellStyle(centerStyle);
			row.createCell(8).setCellValue("Week 3");
			sheetMix.addMergedRegion(new CellRangeAddress(rCountMix-1, rCountMix-1, 8, 10));
			row.getCell(8).setCellStyle(centerStyle);
			row.createCell(11).setCellValue("Week 4");
			sheetMix.addMergedRegion(new CellRangeAddress(rCountMix-1, rCountMix-1, 11, 13));
			row.getCell(11).setCellStyle(centerStyle);
			row.createCell(14).setCellValue("Week 5");
			sheetMix.addMergedRegion(new CellRangeAddress(rCountMix-1, rCountMix-1, 14, 16));
			row.getCell(14).setCellStyle(centerStyle);
			row.createCell(17).setCellValue("Week 6");
			sheetMix.addMergedRegion(new CellRangeAddress(rCountMix-1, rCountMix-1, 17, 19));
			row.getCell(17).setCellStyle(centerStyle);
			row.createCell(20).setCellValue("Week 7");
			sheetMix.addMergedRegion(new CellRangeAddress(rCountMix-1, rCountMix-1, 20, 22));
			row.getCell(20).setCellStyle(centerStyle);
			row.createCell(23).setCellValue("Week 8");
			sheetMix.addMergedRegion(new CellRangeAddress(rCountMix-1, rCountMix-1, 23, 25));
			row.getCell(23).setCellStyle(centerStyle);
			row.createCell(26).setCellValue("Week 9");
			sheetMix.addMergedRegion(new CellRangeAddress(rCountMix-1, rCountMix-1, 26, 28));
			row.getCell(26).setCellStyle(centerStyle);
			row.createCell(29).setCellValue("Week 10");
			sheetMix.addMergedRegion(new CellRangeAddress(rCountMix-1, rCountMix-1, 29, 31));
			row.getCell(29).setCellStyle(centerStyle);
			row.createCell(32).setCellValue("Week 11");
			sheetMix.addMergedRegion(new CellRangeAddress(rCountMix-1, rCountMix-1, 32, 34));
			row.getCell(32).setCellStyle(centerStyle);
			row.createCell(35).setCellValue("Week 12");
			sheetMix.addMergedRegion(new CellRangeAddress(rCountMix-1, rCountMix-1, 35, 37));
			row.getCell(35).setCellStyle(centerStyle);
			row.createCell(38).setCellValue("Week 13");
			sheetMix.addMergedRegion(new CellRangeAddress(rCountMix-1, rCountMix-1, 38, 40));
			row.getCell(38).setCellStyle(centerStyle);
			row.createCell(41).setCellValue("Week 14");
			sheetMix.addMergedRegion(new CellRangeAddress(rCountMix-1, rCountMix-1, 41, 43));
			row.getCell(41).setCellStyle(centerStyle);
			row = sheetMix.createRow(rCountMix);
			rCountMix++;
			for(int wk = 1 ; wk < 15 ; wk++) {
				row.createCell(2 + (wk - 1)*3).setCellValue("Actual wk " + wk);
				if( wk < 3) {
					row.createCell(3 + (wk - 1)*3).setCellValue("Prev. N.A.");
				} else if( wk < 8) {
					row.createCell(3 + (wk - 1)*3).setCellValue("AVG wks 1-" + (wk-2));
				} else {
					row.createCell(3 + (wk - 1)*3).setCellValue("AVG wks" + (wk-6) + "-" + (wk-2));
				}
				row.createCell(4 + (wk - 1)*3).setCellValue("Judged wk" + wk);
			}
		}
		//create SKU column
		for(String k : weeklyMapsSku.get(1).keySet()) { //could pick any of the weeks, using wk 1 here
			row = sheetMix.createRow(rCountMix);
			row.createCell(0).setCellValue(name);
			row.createCell(1).setCellValue(k);
			rCountMix++;
			
		}
		//add mix data
		for (int wk = 1; wk < weeklyMapsSku.size() + 1; wk++) {
			try{
				for(Row r : sheetMix) {
					if (r.getRowNum() > previousLastRowMix) {
						//Actual values
						Cell c = r.createCell(2 + (wk-1)*3);
						if(r.getCell(1).getStringCellValue().contains("PPM")) {
							c.setCellValue(weeklyMapsSku.get(wk).get(r.getCell(1).getStringCellValue()));
							c.setCellStyle(percentageStyle);
							beforeThisRow++;
						}
						
						
						//calculate averages
						c =r.createCell(3 + (wk - 1)*3);
						if(r.getCell(1).getStringCellValue().contains("PPM")) {
							try {
								if(wk < 3) {
									c.setCellValue(0);
								
								} else if(wk < 8) {  //make sure the week wk-2 has the Act uploaded, otherwise use previous AVG available
									if(GTAFunctions.lastActWk >= wk - 2 ) {
										double soma = 0;
										for (int n = 1; n < wk - 1; n++) {
											soma += r.getCell(2 + (n - 1)*3).getNumericCellValue();
										}
										c.setCellValue(soma/(wk - 2));
									} else {
										c.setCellValue(r.getCell(3 + (GTAFunctions.lastActWk + 1)*3).getNumericCellValue());
									}
			
								} else { //make sure the week wk-2 has the Act uploaded, otherwise use previous AVG available
									if(GTAFunctions.lastActWk >= wk - 2 ) {
										double soma = 0;
										for (int n = wk - 6; n < wk - 1; n++) {
											soma += r.getCell(2 + (n - 1)*3).getNumericCellValue();
										}
										c.setCellValue(soma/5);
									} else {
										c.setCellValue(r.getCell(3 + (GTAFunctions.lastActWk + 1)*3).getNumericCellValue());
									}
								
								}
							} catch (IllegalStateException e) {
								c.setCellValue(0);
							}
							c.setCellStyle(percentageStyle);
						}
						
						//Judged values = averages
						c =r.createCell(4 + (wk - 1)*3);
						if(r.getCell(1).getStringCellValue().contains("PPM")) {
							try {
								if(wk < 3) {
									c.setCellValue(0);
								
								} else if(wk < 8) {  //make sure the week wk-2 has the Act uploaded, otherwise use previous AVG available
									if(GTAFunctions.lastActWk >= wk - 2 ) {
										double soma = 0;
										for (int n = 1; n < wk - 1; n++) {
											soma += r.getCell(2 + (n - 1)*3).getNumericCellValue();
										}
										c.setCellValue(soma/(wk - 2));
									} else {
										c.setCellValue(r.getCell(3 + (GTAFunctions.lastActWk + 1)*3).getNumericCellValue());
									}
			
								} else { //make sure the week wk-2 has the Act uploaded, otherwise use previous AVG available
									if(GTAFunctions.lastActWk >= wk - 2 ) {
										double soma = 0;
										for (int n = wk - 6; n < wk - 1; n++) {
											soma += r.getCell(2 + (n - 1)*3).getNumericCellValue();
										}
										c.setCellValue(soma/5);
									} else {
										c.setCellValue(r.getCell(3 + (GTAFunctions.lastActWk + 1)*3).getNumericCellValue());
									}
								
								}
							} catch (IllegalStateException e) {
								c.setCellValue(0);
							}
							c.setCellStyle(percentageStyle2);
						} else {
							if(beforeThisRow > 0) {
								switch (wk) {
								case 1: c.setCellFormula("SUM(E" + (r.getRowNum() + 1 - beforeThisRow) +":E" + r.getRowNum() + ")");
									break;
								case 2: c.setCellFormula("SUM(H" + (r.getRowNum() + 1 - beforeThisRow) +":H" + r.getRowNum() + ")");
									break;
								case 3: c.setCellFormula("SUM(K" + (r.getRowNum() + 1 - beforeThisRow) +":K" + r.getRowNum() + ")");
									break;
								case 4: c.setCellFormula("SUM(N" + (r.getRowNum() + 1 - beforeThisRow) +":N" + r.getRowNum() + ")");
									break;
								case 5: c.setCellFormula("SUM(Q" + (r.getRowNum() + 1 - beforeThisRow) +":Q" + r.getRowNum() + ")");
									break;
								case 6: c.setCellFormula("SUM(T" + (r.getRowNum() + 1 - beforeThisRow) +":T" + r.getRowNum() + ")");
									break;
								case 7: c.setCellFormula("SUM(W" + (r.getRowNum() + 1 - beforeThisRow) +":W" + r.getRowNum() + ")");
									break;
								case 8: c.setCellFormula("SUM(Z" + (r.getRowNum() + 1 - beforeThisRow) +":Z" + r.getRowNum() + ")");
									break;
								case 9: c.setCellFormula("SUM(AC" + (r.getRowNum() + 1 - beforeThisRow) +":AC" + r.getRowNum() + ")");
									break;
								case 10: c.setCellFormula("SUM(AF" + (r.getRowNum() + 1 - beforeThisRow) +":AF" + r.getRowNum() + ")");
									break;
								case 11: c.setCellFormula("SUM(AI" + (r.getRowNum() + 1 - beforeThisRow) +":AI" + r.getRowNum() + ")");
									break;
								case 12: c.setCellFormula("SUM(AL" + (r.getRowNum() + 1 - beforeThisRow) +":AL" + r.getRowNum() + ")");
									break;
								case 13: c.setCellFormula("SUM(AO" + (r.getRowNum() + 1 - beforeThisRow) +":AO" + r.getRowNum() + ")");
									break;
								case 14: c.setCellFormula("SUM(AR" + (r.getRowNum() + 1 - beforeThisRow) +":AR" + r.getRowNum() + ")");
									break;
								}
								beforeThisRow = 0;
							} else {
								c.setCellValue(0);
							}
							
							c.setCellStyle(percentBoldStyle);
						}
						
					}
				}
			} catch (NullPointerException e) {
				//e.printStackTrace();
			} 
		}
		//auto size columns and freeze panes
		for( Cell c : sheetMix.getRow(2)) {
			sheetMix.autoSizeColumn(c.getColumnIndex());
		}
		sheetMix.createFreezePane(2, 2);
				
		//update the previousLastRowMix, subtract 1 since we had added 1 and didn't "use" the row yet
		previousLastRowMix = rCountMix - 1;
		
	}
	
	/**
	 * Forecast by account after CSAM judgment, should be the last (3rd) tab in the file
	 */
	public static void createTemplate(String name, Map<Integer, LinkedHashMap<String, Double>> weeklyMapsSku, boolean isFirst, boolean isLast) {
		Row row;
		int firstRow, lastRow;
		
		//set firstRow, note that we refer to the Forecast By Account sheet (previousLastRow)
		if(isFirst) {
			firstRow = 2;
		} else {
			firstRow = oldPreviousLastRow + 2; //we don't want 0 based here
		}
		
		lastRow = previousLastRow + 1; //we don't want 0 based here
		Cell c;
		
		DataFormat df;
		CellStyle integ;
		//create format styles
		df = wb.createDataFormat();
		integ = wb.createCellStyle();
		integ.setDataFormat(df.getFormat("0"));
		
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
	    if(isFirst) {
	    	row = sheet1.createRow(rCountFore);
	    	rCountFore++;
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
			row.createCell(36).setCellValue("Week 13 JST");
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
			for( Cell c2 : row) {
				sheet1.autoSizeColumn(c2.getColumnIndex());
			}
	    }
		
		//create SKU column
		for(String k : weeklyMapsSku.get(1).keySet()) { //could pick any of the weeks, using wk 1 here
			if( k.contains("PPM")) {
				row = sheet1.createRow(rCountFore);
				row.createCell(0).setCellValue(name);
				row.createCell(1).setCellValue(k);
				RTLMap.put(rCountFore, rowsToLeap);
				rCountFore++;
			} else {
				rowsToLeap++;
			}
		}
		//add mix*subFforecast
		for (int wk = 1; wk < weeklyMapsSku.size() + 1; wk++) {
			try{
				for(Row r : sheet1) {
					if (r.getRowNum() > previousLastRowFore) {
						String subF = GTAFunctions.skuSubF.get(r.getCell(1).getStringCellValue());
						
						switch ( wk) {
						case 1: String formula1 = "Mix!E" + (r.getRowNum()+RTLMap.get(r.getRowNum())) + "*VLOOKUP(\"" + subF + "\", 'Forecast By account'!B" + firstRow + ":P" + lastRow + ", " + (wk + 1) + ", 0)";
							c = r.createCell(10);
							c.setCellFormula(formula1);
							c.setCellStyle(integ);
							break;
						
						case 2: String formula2 = "Mix!H" + (r.getRowNum()+RTLMap.get(r.getRowNum())) + "*VLOOKUP(\"" + subF + "\", 'Forecast By account'!B" + firstRow + ":P" + lastRow + ", " + (wk + 1) + ", 0)";
						c = r.createCell(11);
						c.setCellFormula(formula2);
						c.setCellStyle(integ);
						break;
						
						case 3: String formula3 = "Mix!K" + (r.getRowNum()+RTLMap.get(r.getRowNum())) + "*VLOOKUP(\"" + subF + "\", 'Forecast By account'!B" + firstRow + ":P" + lastRow + ", " + (wk + 1) + ", 0)";
						c = r.createCell(12);
						c.setCellFormula(formula3);
						c.setCellStyle(integ);
						break;
						
						case 4: String formula4 = "Mix!N" + (r.getRowNum()+RTLMap.get(r.getRowNum())) + "*VLOOKUP(\"" + subF + "\", 'Forecast By account'!B" + firstRow + ":P" + lastRow + ", " + (wk + 1) + ", 0)";
						c = r.createCell(13);
						c.setCellFormula(formula4);
						c.setCellStyle(integ);
						break;
						
						case 5: String formula5 = "Mix!Q" + (r.getRowNum()+RTLMap.get(r.getRowNum())) + "*VLOOKUP(\"" + subF + "\", 'Forecast By account'!B" + firstRow + ":P" + lastRow + ", " + (wk + 1) + ", 0)";
						c = r.createCell(14);
						c.setCellFormula(formula5);
						c.setCellStyle(integ);
						break;
						
						case 6: String formula6 = "Mix!T" + (r.getRowNum()+RTLMap.get(r.getRowNum())) + "*VLOOKUP(\"" + subF + "\", 'Forecast By account'!B" + firstRow + ":P" + lastRow + ", " + (wk + 1) + ", 0)";
						c = r.createCell(15);
						c.setCellFormula(formula6);
						c.setCellStyle(integ);
						break;
						
						case 7: String formula7 = "Mix!W" + (r.getRowNum()+RTLMap.get(r.getRowNum())) + "*VLOOKUP(\"" + subF + "\", 'Forecast By account'!B" + firstRow + ":P" + lastRow + ", " + (wk + 1) + ", 0)";
						c = r.createCell(16);
						c.setCellFormula(formula7);
						c.setCellStyle(integ);
						break;
						
						case 8: String formula8 = "Mix!Z" + (r.getRowNum()+RTLMap.get(r.getRowNum())) + "*VLOOKUP(\"" + subF + "\", 'Forecast By account'!B" + firstRow + ":P" + lastRow + ", " + (wk + 1) + ", 0)";
						c = r.createCell(17);
						c.setCellFormula(formula8);
						c.setCellStyle(integ);
						break;
						
						case 9: String formula9 = "Mix!AC" + (r.getRowNum()+RTLMap.get(r.getRowNum())) + "*VLOOKUP(\"" + subF + "\", 'Forecast By account'!B" + firstRow + ":P" + lastRow + ", " + (wk + 1) + ", 0)";
						c = r.createCell(18);
						c.setCellFormula(formula9);
						c.setCellStyle(integ);
						break;
						
						case 10: String formula10 = "Mix!AF" + (r.getRowNum()+RTLMap.get(r.getRowNum())) + "*VLOOKUP(\"" + subF + "\", 'Forecast By account'!B" + firstRow + ":P" + lastRow + ", " + (wk + 1) + ", 0)";
						c = r.createCell(19);
						c.setCellFormula(formula10);
						c.setCellStyle(integ);
						break;
						
						case 11: String formula11 = "Mix!AI" + (r.getRowNum()+RTLMap.get(r.getRowNum())) + "*VLOOKUP(\"" + subF + "\", 'Forecast By account'!B" + firstRow + ":P" + lastRow + ", " + (wk + 1) + ", 0)";
						c = r.createCell(20);
						c.setCellFormula(formula11);
						c.setCellStyle(integ);
						break;
						
						case 12: String formula12 = "Mix!AL" + (r.getRowNum()+RTLMap.get(r.getRowNum())) + "*VLOOKUP(\"" + subF + "\", 'Forecast By account'!B" + firstRow + ":P" + lastRow + ", " + (wk + 1) + ", 0)";
						c = r.createCell(21);
						c.setCellFormula(formula12);
						c.setCellStyle(integ);
						break;
						
						case 13: String formula13 = "Mix!AO" + (r.getRowNum()+RTLMap.get(r.getRowNum())) + "*VLOOKUP(\"" + subF + "\", 'Forecast By account'!B" + firstRow + ":P" + lastRow + ", " + (wk + 1) + ", 0)";
						c = r.createCell(22);
						c.setCellFormula(formula13);
						c.setCellStyle(integ);
						break;
						
						case 14: String formula14 = "Mix!AR" + (r.getRowNum()+RTLMap.get(r.getRowNum())) + "*VLOOKUP(\"" + subF + "\", 'Forecast By account'!B" + firstRow + ":P" + lastRow + ", " + (wk + 1) + ", 0)";
						c = r.createCell(23);
						c.setCellFormula(formula14);
						c.setCellStyle(integ);
						break;
					}
					}
				}
			} catch (NullPointerException e) {
				
			} catch (FormulaParseException e) {
				
			}
		}
		//freeze panes and re-autosize this column after adding content
		sheet1.autoSizeColumn(0);
		sheet1.autoSizeColumn(1);
		sheet1.createFreezePane(2, 1, 8, 1);
		
		//update the previousLastRowMix, subtract 1 since we had added 1 and didn't "use" the row yet
		previousLastRowFore = rCountFore - 1;
		
		//create and save the xls file
		if(isLast) {
			wb.setActiveSheet(1);
			try {
				fileOut = new FileOutputStream(GTAGUI.inputPath.getParent() + "/ForecastMix_" + String.valueOf((todayNow.get(Calendar.MONTH)+1)) + String.valueOf(todayNow.get(Calendar.DAY_OF_MONTH)) + 
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
	}
	
}
