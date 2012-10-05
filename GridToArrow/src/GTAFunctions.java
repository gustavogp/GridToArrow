import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;


public class GTAFunctions {
	
	public static void countAndCheck(File inputPath) {
		ArrayList<File> folderFiles = new ArrayList<File>();
		File geo = null;
		
		for (File f : inputPath.listFiles()){
			folderFiles.add(f);
		}
		
		GTAGUI.generalMessage("Found " + (folderFiles.size()-1) + " files in folder " + inputPath.getName());
		for (File f2 : folderFiles) {
			if (!(f2.getName().contains("DS_Store"))){
				GTAGUI.generalMessage(f2.getName() + ", Last Modified on " + (new Date(f2.lastModified())).toString());
			}
			if (f2.getName().contains("geo") || f2.getName().contains("GEO")) {
				geo = f2;
			}
		}
		if (geo == null) {
			GTAGUI.generalMessage("No \"geo\" file was found");
		} else {
			calculateMix(geo);
		}
	}
	
	public static void calculateMix(File geo){
		Workbook wb = new HSSFWorkbook();
		Sheet sheet1 = null, sheet0 = null;
		FileInputStream readStr = null;
		
		try {
			readStr = new FileInputStream(geo);
			wb = new HSSFWorkbook(readStr);
			sheet0 = wb.getSheet("Product");
			sheet1 = wb.getSheet("PPN");
			
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		
		Row wkRow0 = sheet0.getRow(7);
		int wks = 0;
		
		//how many weeks in this quarter
		wks = calcWeeks(wkRow0);
		GTAGUI.generalMessage("Weeks in this quarter: " + wks);
		
		// list subfamilies
		ArrayList<String> subFamilies = new ArrayList<String>();
		for (Row r : sheet0) {
			try {
				if (r.getRowNum() > 6 && 
					!(r.getCell(1).getStringCellValue().contains("Total")) &&
					!(r.getCell(1).getCellType()==Cell.CELL_TYPE_BLANK) && 
					!(sheet0.getRow(r.getRowNum()+1).getCell(2).getCellType()==Cell.CELL_TYPE_STRING)) {
				subFamilies.add(r.getCell(1).getStringCellValue());
				}
			} catch (NullPointerException e) {
				//do nothing
			} catch (IllegalStateException e) {
				//do nothing
			}
		}
		GTAGUI.generalMessage("SubFamilies: " + subFamilies); //testing only, delete later
		
		//create total/wk/subfamily
		Map<Integer, HashMap<String, Integer>> weeklyMapsSF = new HashMap<Integer, HashMap<String, Integer>>();
		weeklyMapsSF = createSubFPerWeek( wks, wkRow0, sheet0);
		
		//create mix/wk/sku
		Map<Integer, HashMap<String, Integer>> weeklyMapsSku = new HashMap<Integer, HashMap<String, Integer>>();
		weeklyMapsSku = createMixPerWeek(wks, sheet1);
		GTAGUI.generalMessage("Total Quantities per Sku per week" + weeklyMapsSku);//testing only, delete later
		
		//create map Sku to SubFamily
		
	}
	
	public static int calcWeeks (Row wkRow0) {
		int nwks = 0;
		for (Cell c : wkRow0) {
			try {
				if (c.getStringCellValue().equalsIgnoreCase("ST")) {
					nwks = (int) wkRow0.getCell((c.getColumnIndex() - 1)).getNumericCellValue();
					break;
				}
			} catch (NullPointerException e) {
				//e.printStackTrace();
			} catch (IllegalStateException e) {
				//do nothing
			}
		}
		return nwks;
	}
	
	public static Map<Integer, HashMap<String, Integer>> createSubFPerWeek(int wks, Row wkRow0, Sheet sheet0 ){
		Map<Integer, HashMap<String, Integer>> weeklyMapsSF = new HashMap<Integer, HashMap<String, Integer>>();
		int columnIndexSF = 0;
		for (int w = 1; w < wks + 1; w++) {
			Map<String, Integer> subFPerWk = new HashMap<String, Integer>();//create a new map for each wk
			//find the column index of this wk
			for (Cell d : wkRow0) {
				try {
					if ((new Double(d.getNumericCellValue())).intValue() == w) {
						columnIndexSF = d.getColumnIndex();
						break;
					}
				} catch (NullPointerException e) {
				//	e.printStackTrace();
				} catch (IllegalStateException e) {
				//	e.printStackTrace();
				}
			}
			//build subFamily Per Week map
			for (Row rw : sheet0) {
				try {
					if (rw.getRowNum() > 6 && 
							!(rw.getCell(1).getCellType()==Cell.CELL_TYPE_BLANK) &&
							!(rw.getCell(1).getStringCellValue().contains("Total")) && 
							!(sheet0.getRow(rw.getRowNum()+1).getCell(2).getCellType()==Cell.CELL_TYPE_STRING)) {
						subFPerWk.put(rw.getCell(1).getStringCellValue(), (new Double(rw.getCell(columnIndexSF).getNumericCellValue())).intValue());
					}
				} catch (NullPointerException e) {
				//	e.printStackTrace();
				} catch (IllegalStateException e) {
				//	e.printStackTrace();
				}
				
			}
			weeklyMapsSF.put(w, (HashMap<String, Integer>) subFPerWk);
		}
		return weeklyMapsSF;
	}
	
	public static Map<Integer, HashMap<String, Integer>> createMixPerWeek(int wks, Sheet sheet1 ){
		Row wkRow1 = sheet1.getRow(7);
		Map<Integer, HashMap<String, Integer>> weeklyMapsSku = new HashMap<Integer, HashMap<String, Integer>>();
		int columnIndexSku = 0;
		for (int w = 1; w < wks + 1; w++) {
			Map<String, Integer> skuPerWk = new HashMap<String, Integer>();//create a new map for each wk
			//find the column index of this wk
			for (Cell d : wkRow1) {
				try {
					if ((new Double(d.getNumericCellValue())).intValue() == w) {
						columnIndexSku = d.getColumnIndex();
						break;
					}
				} catch (NullPointerException e) {
				//	e.printStackTrace();
				} catch (IllegalStateException e) {
				//	e.printStackTrace();
				}
			}
			//build Sku Per Week map
			for (Row rw : sheet1) {
				try {
					if (rw.getRowNum() > 6 && 
							!(rw.getCell(3).getCellType()==Cell.CELL_TYPE_BLANK) &&
							rw.getCell(3).getStringCellValue().contains("PPM"))  {
						skuPerWk.put(rw.getCell(3).getStringCellValue(), (new Double(rw.getCell(columnIndexSku).getNumericCellValue())).intValue());
					}
				} catch (NullPointerException e) {
				//	e.printStackTrace();
				} catch (IllegalStateException e) {
				//	e.printStackTrace();
				}
				
			}
			weeklyMapsSku.put(w, (HashMap<String, Integer>) skuPerWk);
		}
		return weeklyMapsSku;
	}
	
	public static Map<String, String> createSkuSubF() {
		Map<String, String> skuSubF = new HashMap<String, String>();
		
		return skuSubF;
	}
}
