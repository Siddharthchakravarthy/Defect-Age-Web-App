package net.JavaGS;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.time.LocalDate;
import java.util.Arrays;
import java.util.Iterator;
import java.util.List;

import javax.servlet.http.Part;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelColumnFilter {
	 public void filterExcel(FileInputStream fileContent) {		 	
		 	LocalDate currentDate = LocalDate.now();
//	        String inputFile = "C:\\Users\\PC-1\\eclipse-workspace\\TEst\\target\\Octane_defects_filtered_27_6_2023_11_23_08_am.xlsx";
	        String outputFile = "C:\\DEV\\UploadedFile\\"+currentDate+".xlsx";
	        String sheetName = "Data";
	        int filterColumnIndex = 6; // Assuming the filter column is the third column (0-based index)
	        List<String> filterValues = Arrays.asList("Deferred", "Fixed", "In Progress", "New", "Review"); // Filter values
	        System.out.println("Sidd");
	        try (XSSFWorkbook workbook = new XSSFWorkbook(fileContent)) {
	        	
	            XSSFSheet sheet = workbook.getSheet(sheetName);

	            // Iterate through the rows and apply the filter
	            Iterator<Row> rowIterator = sheet.iterator();
	            while (rowIterator.hasNext()) {
	                Row row = rowIterator.next();
	                Cell cell = row.getCell(filterColumnIndex);

	                // Check if it is the header row
	                if (row.getRowNum() == 0) {
	                    //Un-hide the header row
	                    row.setZeroHeight(false);
	                } else {
	                    if (cell != null && filterValues.contains(cell.getStringCellValue())) {
	                        // Keep the row
	                        row.setZeroHeight(false);
	                    } else {
	                        // Hide the row
	                        row.setZeroHeight(true);
	                    }
	                }
	            }
	            
	            // Save the filtered data to a new workbook
	            try (FileOutputStream fos = new FileOutputStream(outputFile)) {
	                workbook.write(fos);
	                
	                System.out.println("Filtered data saved successfully to " + outputFile);
	            }
	        } catch (IOException e) {
	            e.printStackTrace();
	        }
	    }
}
