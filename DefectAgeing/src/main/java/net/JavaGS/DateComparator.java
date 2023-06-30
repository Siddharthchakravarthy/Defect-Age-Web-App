package net.JavaGS;


import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.time.Instant;
import java.time.LocalDate;
import java.time.ZoneId;
import java.time.temporal.ChronoUnit;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class DateComparator {
	public static String filePath;
	public void PrintCurrentDate(FileInputStream fileContent) {   
				LocalDate currentDate = LocalDate.now();
	            filePath = "C:\\DEV\\UploadedFile\\"+currentDate+".xlsx";
	            try (
	            		Workbook workbook = WorkbookFactory.create(fileContent)) {
	                Sheet sheet = workbook.getSheetAt(2); // Assuming you're working with the first sheet

	                for (Row row : sheet) {
	                	if(row.getRowNum() == 0) {
	                		Cell cell1 = row.createCell(14);
	                		cell1.setCellValue("Current Date");
	                		continue;
	                	}
	                	if(row.getRowNum() == sheet.getLastRowNum()) {
	                		continue;
	                	}
	                	CellStyle cellStyle = workbook.createCellStyle();
	                	CreationHelper createHelper = workbook.getCreationHelper();
	                	cellStyle.setDataFormat(
	                	    createHelper.createDataFormat().getFormat("dd-mm-yy"));
	                	Cell cell = row.createCell(14);
	                	cell.setCellValue(new Date());
	                	cell.setCellStyle(cellStyle);
	                }
		               
//	                // Save the modified workbook
	                try (FileOutputStream outputStream = new FileOutputStream(filePath)) {
	                    workbook.write(outputStream);
	                }
	            } catch (IOException e) {
	                e.printStackTrace();
	            }
	        }
	
	public void getDateDifference() {
		LocalDate currentDate = LocalDate.now();
        filePath = "C:\\DEV\\UploadedFile\\"+currentDate+".xlsx";
        try (Workbook workbook = WorkbookFactory.create(new FileInputStream(filePath))) {
            Sheet sheet = workbook.getSheetAt(2); // Assuming you're working with the first sheet

            for (Row row : sheet) {
            	if(row.getRowNum() == 0) {
            		Cell cell1 = row.createCell(15);
            		cell1.setCellValue("Date Difference");
            		continue;
            	}
                Cell date1Cell = row.getCell(2); // Column 2
                Cell date2Cell = row.getCell(14); // Column 14
                Cell dateDiffCell = row.createCell(15); // Column 15

                if (date1Cell != null && date1Cell.getCellType() == CellType.NUMERIC &&
                      date2Cell != null && date2Cell.getCellType() == CellType.NUMERIC) {
                	  double date1Value = date1Cell.getNumericCellValue();
                      double date2Value = date2Cell.getNumericCellValue();

                      java.util.Date date1 = DateUtil.getJavaDate(date1Value);
                      java.util.Date date2 = DateUtil.getJavaDate(date2Value);

                      LocalDate localDate1 = Instant.ofEpochMilli(date1.getTime()).atZone(ZoneId.systemDefault()).toLocalDate();
                      LocalDate localDate2 = Instant.ofEpochMilli(date2.getTime()).atZone(ZoneId.systemDefault()).toLocalDate();

                    
                      long daysDiff = ChronoUnit.DAYS.between(localDate1, localDate2);

                    dateDiffCell.setCellValue(daysDiff);
                }
            }
            System.out.println("Sidd");
            // Save the modified workbook
            try (FileOutputStream outputStream = new FileOutputStream(filePath)) {
                workbook.write(outputStream);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
