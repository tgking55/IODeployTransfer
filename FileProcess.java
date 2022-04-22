package IODeployTransfer;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;




public class FileProcess {
	private static final Logger log = LogManager.getLogger(FileProcess.class);
	private boolean isNoOrg  = false;
	Map<String, Map> readExcel(File excel) throws IOException{		
		Map<String, Map> table = new HashMap(); //table<ip, map<分表id, row arraylist>>					
		Workbook wb = null;
		log.info("readExcel ...."+excel.getName());
		try {
			wb = WorkbookFactory.create(excel);					
			Sheet sheet = wb.getSheetAt(0);	
			if(sheet == null)
				return null;
			int rowNum = sheet.getLastRowNum() + 1;				
			Map<String, ArrayList> plcMap = new HashMap();
			String ip="";
			boolean is_field_name = false;
			for(int i =0; i< rowNum ; i++) {
				//略過欄位名稱那行
				if(is_field_name) {
					is_field_name = false;
					continue;
				}					
				Row row = sheet.getRow(i);	
				if(row == null) {
					log.error("row == null");
					break;
				}
				int cnum = row.getPhysicalNumberOfCells();
				String key="";
				boolean empty = true;
				for (int c = 0; c < cnum; c++) {
					 if(c > 8) //第9欄以後跳過
						break;
					Cell cell = row.getCell(c);		
					if(cell != null) {
						cell.setCellType(CellType.STRING);
						String c_str = cell.toString().trim();
						c_str=c_str.replace(String.valueOf((char)160)," ");						
						c_str=c_str.trim();
//						System.out.println("xxxxxxx"+c_str+"xxx");
						
											
						if(c == 0 && c_str=="")//下一個PLC
							break;

						
						
						
						int get_ip = c_str.indexOf("192.168.50.");
						if(get_ip > 0) {
							ip = c_str.substring(get_ip, c_str.length());
							log.info(ip);
							plcMap = new HashMap();
							table.put(ip, plcMap);
							is_field_name = true; //略過欄位名稱那行
							break;
						}else {	
							plcMap = table.get(ip);	 
							c_str = c_str.trim();
							switch (c) {
								case 0:											
									int iof = c_str.indexOf('.');
									if(iof > 0) {
										c_str = c_str.substring(0, iof);
									}										
									key = c_str;
									plcMap.put(key, new ArrayList());
									break;
								case 1:	case 3: case 4: case 5: case 6: case 7: case 8:										
									if(c_str.trim() != "")
										empty = false;
									if(key.trim() == "") {
										log.error("key = "+key);
										break; 
									}
										
									plcMap.get(key).add(c_str);												
									table.put(ip, plcMap);										
									break;
								case 2:  //樓	
									ArrayList ary = plcMap.get(key);	
									if(c_str.equals("B1F "))
										System.out.println();
									if(c_str !="") 										
										c_str = c_str.trim()+"F";
									
									
									ary.add(c_str);	
									log.info(c_str);
									plcMap.put(key, ary);
									table.put(ip, plcMap);	
									break;
								default:
									break;
							}//end switch (c)
						}//end if(get_ip > 0) 						
					}//end if(row.getCell(c) != null)						
				}//end for each cell	
				
				//刪掉空的分表
				if(!is_field_name && empty) {
					Map<String, ArrayList> plc = table.get(ip);
					if(plc.containsKey(key)) {
						plc.remove(key);
						table.put(ip, plc);
					}					
				}	
			
			}//end for each row
			
			//刪掉空的PLC
			Map removeMap = new HashMap();
			removeMap.putAll(table);				
			for (Map.Entry<String, Map> entry : table.entrySet()) {				
				String key = entry.getKey();
				if(table.get(key).size() == 0)						
					removeMap.remove(key);
			}
			table.clear();
			table.putAll(removeMap);
			wb.close();	
			log.info("end of readExcel ....");
		} catch (IOException e) {	
			if(wb != null)
				wb.close();
			e.printStackTrace();
			log.error(e.getMessage());
		}
		return table;

	}
	
	
	boolean writeToExcel(Map<String, Map> table, String path){
		log.info("writeToExcel ....");
        Workbook wb = new XSSFWorkbook();  
        Sheet sheet =  sheet = wb.createSheet("配置檔");   
        Row row = null;
        Cell cell = null;
        boolean isMore3 = false;   

        Iterator<Map.Entry<String, Map>> itr = table.entrySet().iterator(); 
        while(itr.hasNext())
        {
             Map.Entry<String, Map> entry = itr.next();
             String[] sub_str = entry.getKey().split("\\.");
             if(sub_str == null || sub_str.length == 0)
            	 break;

             int ip = Integer.valueOf(sub_str[sub_str.length-1].substring(1,3)); //ip 後二位 101 ==> 1
             int rbegin = (ip-1)*7+13;
             int rend = rbegin+7;
             //先create row
             for(int r = rbegin; r < rend; r++) {
            	 row = sheet.createRow(r);
            	 if(r == rbegin) {
            		 cell = row.createCell(0);
            		 cell.setCellValue("PLC");
            		 cell = row.createCell(1);
            		 cell.setCellValue(ip+3);
            		 
            		 cell = row.createCell(14);
            		 cell.setCellValue(entry.getKey()); //PLC IP
            		 
            		 
            	 }else if(r == rbegin+4) {
            		 cell = row.createCell(1);
            		 cell.setCellValue("PLC");
            	 }else if(r == rbegin+5) {
            		 cell = row.createCell(1);
            		 cell.setCellValue("遠端控制");
            	 }else if(r == rbegin+6) {
            		 cell = row.createCell(1);
            		 cell.setCellValue("手動設定區");
            	 }
             }
             
        Map<String, ArrayList> plcMap = entry.getValue();
   		for (Map.Entry<String, ArrayList> rowMap : plcMap.entrySet()) {
   			int meter_id = Integer.valueOf(rowMap.getKey());
   			int col = meter_id +1;
   			
   			ArrayList<String> rowlist = rowMap.getValue();   		
   			
   			//分表ID
   			row = sheet.getRow(rbegin);
   			cell = row.createCell(col);
   			cell.setCellValue(String.valueOf(meter_id));
   			
   			switch(meter_id) {
   			case 9:
   				row = sheet.getRow(rbegin+1);
   	  	   		cell = row.createCell(col);
   				cell.setCellValue("校園總電表");
   				setTotalMeter(sheet, row, cell, rowlist, rbegin, col);
   				break;
   			case 10:
   				row = sheet.getRow(rbegin+1);
   	  	   		cell = row.createCell(col);
   				cell.setCellValue("冷氣總電表");
   				setTotalMeter(sheet, row, cell, rowlist, rbegin, col);
   				break;
   			case 11:
   				row = sheet.getRow(rbegin+1);
   	  	   		cell = row.createCell(col);
   				cell.setCellValue("PV電表");
   				setTotalMeter(sheet, row, cell, rowlist, rbegin, col);
   				break;
   			default:  
   				log.info("rowlist = "+rowlist);
   				System.out.println("rowlist.get(6).trim()= "+rowlist.get(6).trim());
   				int airNum = Integer.valueOf(rowlist.get(6).trim()); 
   				if(airNum > 3)  //超過3台冷氣
   					isMore3 = true;
   				   				
   				int air_begin = (meter_id - 1)*3+1; //冷氣ID 開始
   				int air_end = (meter_id - 1)*3+airNum; //冷氣ID 結束
   				int r1 = 1;
   				for(int i = air_begin; i <= air_end; i++) {
   					row = sheet.getRow(rbegin+r1);
   	   	  	   		cell = row.createCell(col);
   	   	  	   		cell.setCellValue(String.valueOf(i)); //冷氣ID
   	   	  	   		r1++;
   				}
   				
   				//大樓、樓層、區域
				row = sheet.getRow(rbegin+4);
	  	   		cell = row.createCell(col);
	  	   		String building = rowlist.get(0).trim();
				cell.setCellValue(rowlist.get(0).trim());
				
				row = sheet.getRow(rbegin+5);
	  	   		cell = row.createCell(col);
	  	   		String floor = rowlist.get(1).trim();
				cell.setCellValue(floor);
				
				row = sheet.getRow(rbegin+6);
	  	   		cell = row.createCell(col);
	  	   		String block = rowlist.get(2).trim();
				cell.setCellValue(block);
				
				
				
				if(building == "" || floor == "" || block=="") {
					isNoOrg = true;
					log.error(rowlist+" no Org =============================================================");
				}
   				
   				break;
   			}
   		}
       }
   
        try {
        	if(isNoOrg) {
        		log.error("NO ORG");        		
        	}else
        		outputFile(wb, path); 
        } catch (IOException e) {
		    e.printStackTrace();
		    log.error(e.getMessage());
		}
        return isMore3;
	}
	
	//設定總電表 大樓、樓層、區域
	void setTotalMeter(Sheet sheet, Row row, Cell cell, ArrayList<String> rowlist, int r, int c) {		
		row = sheet.getRow(r+4);
	   	cell = row.createCell(c);
		cell.setCellValue("總電表");
		
		row = sheet.getRow(r+5);
	   	cell = row.createCell(c);
	   	String building = rowlist.get(0).trim();
	   	String floor = rowlist.get(1).trim();
	   	cell.setCellValue(rowlist.get(0).trim()+" "+rowlist.get(1).trim()); //大樓 樓層
		
		
		row = sheet.getRow(r+6);
	   	cell = row.createCell(c);
	   	String block = rowlist.get(2).trim();
		cell.setCellValue(rowlist.get(2).trim()); //區域
		
		
		if(building == "" || floor == "" || block=="") {
			isNoOrg = true;	
			log.error(rowlist+" no Org =============================================================");
		}
		
		
		
	}
	

	
	public void outputFile(Workbook wb, String path)  throws IOException{
		try {
			FileOutputStream fos = new FileOutputStream(new File(path));		    
		    wb.write(fos);
		    wb.close();
		    fos.flush();
			fos.close();
		} catch (IOException e) {
		    e.printStackTrace();
		    log.error(e.getMessage());
		}
	}
	
	public void moveFile(File excel, String fileName, Map<String, Map> table) throws IOException {
		String path = "./outDir/學校設備配置表-" + fileName + ".xlsx";
		try {
			boolean isMore3 = false; //超過3台冷氣
			isMore3 = writeToExcel(table, path);			
			if(isNoOrg) {
				isNoOrg = false;				
				String more3Path="./noOrg/" + fileName+ ".xlsx";
				Files.move(Paths.get(excel.getAbsolutePath()), Paths.get(more3Path),  StandardCopyOption.REPLACE_EXISTING);				
			}else if(isMore3) {
				log.info(fileName+" more than 3");
				String more3Path="./More3/" + fileName+ ".xlsx";
				Files.move(Paths.get(excel.getAbsolutePath()), Paths.get(more3Path),  StandardCopyOption.REPLACE_EXISTING);
			}else {
				String OK = "./OK/" + fileName + ".xlsx";
				Files.move(Paths.get(excel.getAbsolutePath()), Paths.get(OK),  StandardCopyOption.REPLACE_EXISTING);
			}			
			
			
			log.info("OK");
		} catch(Exception e) {				
			e.printStackTrace();
			log.error(e.getMessage());
			String notOK = "./notOK/" +fileName+ ".xlsx";
			Files.move(Paths.get(excel.getAbsolutePath()), Paths.get(notOK),  StandardCopyOption.REPLACE_EXISTING);
			log.info("notOK");
		}
		
	}
	
	
}
