package IODeployTransfer;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.util.Map;

import org.apache.commons.io.FilenameUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.Workbook;

public class IODeployTransfer {
	private static final Logger log = LogManager.getLogger(IODeployTransfer.class);
	
	public static void main(String[] args) throws IOException {
		FileProcess fp = new FileProcess();
		File dir = new File("./inDir");			
		for (File excel : dir.listFiles()) {
			String ext= FilenameUtils.getExtension(excel.getName());
			String fileName = FilenameUtils.getName(excel.getName()).replace("." + ext, "");
			
			Map<String, Map> table = fp.readExcel(excel); 
			if (table == null)
				System.exit(0);
			String path = "./outDir/學校設備配置表-" + fileName + ".xlsx";
			fp.moveFile(excel, fileName,table);
			
			/*
			
			try {
				boolean isMore3 = false; //超過3台冷氣
				isMore3 = fp.writeToExcel(table, path);
				String OK = "./OK/" + excel.getName();
				if(isMore3) {
					log.info(fileName+" more than 3");
					String more3Path="./More3/" + excel.getName();
//					Files.move(Paths.get(excel.getAbsolutePath()), Paths.get(more3Path),  StandardCopyOption.REPLACE_EXISTING);
				}else {
//					Files.move(Paths.get(excel.getAbsolutePath()), Paths.get(OK),  StandardCopyOption.REPLACE_EXISTING);
				}
				
				
				
				log.info("OK");
			} catch(Exception e) {				
				e.printStackTrace();
				log.error(e.getMessage());
				String notOK = "./notOK/" + excel.getName();
//				Files.move(Paths.get(excel.getAbsolutePath()), Paths.get(notOK),  StandardCopyOption.REPLACE_EXISTING);
				log.info("notOK");
			}
			
			
			*/
			
		}
		
		log.info("end of IOTableTransfer ..................");
	}

}
