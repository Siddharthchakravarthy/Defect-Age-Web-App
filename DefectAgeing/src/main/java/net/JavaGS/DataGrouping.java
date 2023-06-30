package net.JavaGS;


import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class DataGrouping {
	 public void createSummary(FileInputStream FileContent, String Path) {
		 String sheetName = "Data";
		 int count1 = 0, count2 = 0, count3 =0, count4 =0, count5 =0, count6 =0, count7 =0, count8 =0, count9 =0, count10 =0, count11 =0, count12 =0, count13 =0, count14 =0, count15 =0, count16 =0, count17 =0, count18 =0, count19 = 0, count20 = 0;
		 
		 try (XSSFWorkbook workbook = new XSSFWorkbook(FileContent)){
			 	
	            XSSFSheet sheet = workbook.getSheet(sheetName);
	            XSSFSheet newSheet = workbook.createSheet("Summary Report");
	            //First Row Headers
	            Row row1 = newSheet.createRow(0);
	            Cell cell01 = row1.createCell(0);
	            
	            
	            cell01.setCellValue("Aging");
	            Cell cell02 = row1.createCell(1);
	            cell02.setCellValue("Urgent");
	            Cell cell03 = row1.createCell(2);
	            cell03.setCellValue("High");
	            Cell cell04 = row1.createCell(3);
	            cell04.setCellValue("Low");
	            Cell cell05 = row1.createCell(4);
	            cell05.setCellValue("Medium");
	            Cell cell06 = row1.createCell(5);
	            cell06.setCellValue("Very High");
	            Cell cell07 = row1.createCell(6);
	            cell07.setCellValue("Grand Total");
	            
	            //First Column 
	            Row row2 = newSheet.createRow(1);
	            Cell cell1 = row2.createCell(0);
	            cell1.setCellValue("< 7");
	            Row row3 = newSheet.createRow(2);
	            Cell cell2 = row3.createCell(0);
	            cell2.setCellValue("8 - 15");
	            Row row4 = newSheet.createRow(3);
	            Cell cell3 = row4.createCell(0);
	            cell3.setCellValue("16 - 30");
	            Row row5 = newSheet.createRow(4);
	            Cell cell4 = row5.createCell(0);
	            cell4.setCellValue("> 30");
	            Row row06 = newSheet.createRow(5);
	            Cell cell5 = row06.createCell(0);
	            cell5.setCellValue("Grand Total");
	            
	            //Data Sheet 
	            for (Row row : sheet) {
	            	if(row.getRowNum() == 0) {
	            		continue;
	            	}
	            	Cell phaseCell = row.getCell(6);
	            	if(phaseCell != null) {
		            	if(phaseCell.getStringCellValue().equalsIgnoreCase("Closed")){
		            		continue;
		            	}
		            	else {
			            	Cell cell = row.getCell(15);
			            	
//			            	if(cell == null) {
//			            		cell = row.createCell(15);
//			            	}
//			            	
//			            	System.out.println(cell.getStringCellValue());
			            	
			            	double age = cell.getNumericCellValue();
			            	
			            	Cell priorityCell = row.getCell(4);
			            	String priority = priorityCell.getStringCellValue();
			            	
			            	if(priority.equalsIgnoreCase("Urgent")) {
				            	if(age <= 7 ) {
				            		count1++;
				            		Cell cell11 = row2.createCell(1);
				            		cell11.setCellValue(count1);
				            	}
								if(age >= 8 & age <= 15) {
									count2++;
				            		Cell cell21 = row3.createCell(1);
				            		cell21.setCellValue(count2);           		
								}
								if(age > 15 & age <=30) {
									count3++;
				            		Cell cell31 = row4.createCell(1);
				            		cell31.setCellValue(count3);
								}
								if(age > 30) {
									count4++;
				            		Cell cell41 = row5.createCell(1);
				            		cell41.setCellValue(count4);
								}
			            	}
			            	else if(priority.equalsIgnoreCase("high")) {
				            	if(age <= 7 ) {
				            		count5++;
				            		Cell cell21 = row2.createCell(2);
				            		cell21.setCellValue(count5);
				            	}
								if(age >= 8 & age <= 15) {
									count6++;
				            		Cell cell22 = row3.createCell(2);
				            		cell22.setCellValue(count6);           		
								}
								if(age > 15 & age <=30) {
									count7++;
				            		Cell cell32 = row4.createCell(2);
				            		cell32.setCellValue(count7);
								}
								if(age > 30) {
									count8++;
				            		Cell cell42 = row5.createCell(2);
				            		cell42.setCellValue(count8);
								}
			            	}
			            	else if(priority.equalsIgnoreCase("Low")) {
				            	if(age <= 7 ) {
				            		count9++;
				            		Cell cell13 = row2.createCell(3);
				            		cell13.setCellValue(count9);
				            	}
								if(age >= 8 & age <= 15) {
									count10++;
				            		Cell cell23 = row3.createCell(3);
				            		cell23.setCellValue(count10);           		
								}
								if(age > 15 & age <=30) {
									count11++;
				            		Cell cell33 = row4.createCell(3);
				            		cell33.setCellValue(count11);
								}
								if(age > 30) {
									count12++;
				            		Cell cell43 = row5.createCell(3);
				            		cell43.setCellValue(count12);
								}
			            	}
			            	else if(priority.equalsIgnoreCase("Medium")) {
				            	if(age <= 7 ) {
				            		count13++;
				            		Cell cell14 = row2.createCell(4);
				            		cell14.setCellValue(count13);
				            	}
								if(age >= 8 & age <= 15) {
									count14++;
				            		Cell cell24 = row3.createCell(4);
				            		cell24.setCellValue(count14);           		
								}
								if(age > 15 & age <=30) {
									count15++;
				            		Cell cell34 = row4.createCell(4);
				            		cell34.setCellValue(count15);
								}
								if(age > 30) {
									count16++;
				            		Cell cell44 = row5.createCell(4);
				            		cell44.setCellValue(count16);
								}
			            	}
			            	else if(priority.equalsIgnoreCase("Very high")) {
				            	if(age <= 7 ) {
				            		count17++;
				            		Cell cell15 = row2.createCell(5);
				            		cell15.setCellValue(count17);
				            	}
								if(age >= 8 & age <= 15) {
									count18++;
				            		Cell cell25 = row3.createCell(5);
				            		cell25.setCellValue(count18);            		
								}
								if(age > 15 & age <=30) {
									count19++;
				            		Cell cell35 = row4.createCell(5);
				            		cell35.setCellValue(count19);
								}
								if(age > 30) {
									count20++;
				            		Cell cell45 = row5.createCell(5);
				            		cell45.setCellValue(count20);
								}
		            	}
		            	}
		            	//Row wise Grand Total
			            Cell cell51 = row06.createCell(1);
			            cell51.setCellValue(count1 + count2 + count3 + count4);
			            
			            Cell cell52 = row06.createCell(2);
			            cell52.setCellValue(count5 + count6 + count7 + count8);
			            
			            Cell cell53 = row06.createCell(3);
			            cell53.setCellValue(count9 + count10 + count11 + count12);
			            
			            Cell cell54 = row06.createCell(4);
			            cell54.setCellValue(count13 + count14 + count15 + count16);
			            
			            Cell cell55 = row06.createCell(5);
			            cell55.setCellValue(count17 + count18 + count19 + count20);
			            
			            Cell cell56 = row06.createCell(6);
			            cell56.setCellValue(count1 + count2 + count3 + count4 + count5 + count6 + count7 + count8 + count9 + count10 + count11 + count12 + count13 + count14 + count15 + count16 + count17 + count18 + count19 + count20);
		            
			            //Column wise Grand Total
			            Cell cell16 = row2.createCell(6);
			            cell16.setCellValue(count1 + count5 + count9 + count13 + count17);
			            
			            Cell cell26 = row3.createCell(6);
			            cell26.setCellValue(count2 + count6 + count10 + count14 + count18);
			            
			            Cell cell36 = row4.createCell(6);
			            cell36.setCellValue(count3 + count7 + count11 + count15 + count19);
			            
			            Cell cell46 = row5.createCell(6);
			            cell46.setCellValue(count4 + count8 + count12 + count16 + count20);
		            }
	            }
	            try (FileOutputStream fos = new FileOutputStream(Path)) {
	                workbook.write(fos);
	                workbook.close();
	                System.out.println("Filtered data saved successfully to " + Path);
	            }
	        } catch (IOException e) {
	            e.printStackTrace();
	        }
	 }
}
