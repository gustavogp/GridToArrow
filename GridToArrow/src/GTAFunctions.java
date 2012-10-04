import java.io.File;
import java.util.ArrayList;
import java.util.Date;


public class GTAFunctions {
	
	public static void countAndCheck(File inputPath) {
		ArrayList<File> folderFiles = new ArrayList<File>();
		
		for (File f : inputPath.listFiles()){
			folderFiles.add(f);
		}
		
		GTAGUI.generalMessage("Found " + (folderFiles.size()-1) + " files in folder " + inputPath.getName());
		for (File f2 : folderFiles) {
			if (!(f2.getName().contains("DS_Store"))){
				GTAGUI.generalMessage(f2.getName() + ", Last Modified on " + (new Date(f2.lastModified())).toString());

			}
		}
		
		
		
		
	}
	
}
