import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;

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
		
		//how many weeks in this quarter
		Row wkRow = sheet0.getRow(7);
		int wks = 0;
		for (Cell c : wkRow) {
			try {
				if (c.getStringCellValue().equalsIgnoreCase("ST")) {
					wks = (int) wkRow.getCell((c.getColumnIndex() - 1)).getNumericCellValue();
					break;
				}
			} catch (NullPointerException e) {
				e.printStackTrace();
			} catch (IllegalStateException e) {
				//do nothing
			}
		}
		GTAGUI.generalMessage("Weeks in this quarter: " + wks);
		
	}
}
