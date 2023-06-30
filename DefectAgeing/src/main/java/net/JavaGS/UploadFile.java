package net.JavaGS;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.PrintWriter;
import java.time.LocalDate;
import java.util.Random;

import javax.servlet.ServletException;
import javax.servlet.annotation.MultipartConfig;
import javax.servlet.annotation.WebServlet;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.servlet.http.Part;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.nodes.Entities.EscapeMode;


@WebServlet("/UploadFile")
@MultipartConfig(
		location = "C:\\DEV\\UploadedFile", 
		fileSizeThreshold = 1024 * 1024,
		maxFileSize = 1024 * 1024 * 10,
		maxRequestSize = 1024 * 1024 * 11
)
public class UploadFile extends HttpServlet {
	private static final long serialVersionUID = 1L;
	private static boolean runOnce = true;
    public UploadFile() {
        super();
        // TODO Auto-generated constructor stub
    }
    
    protected void doGet(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
        response.setHeader("Cache-Control", "no-cache, no-store, must-revalidate");
        response.setHeader("Pragma", "no-cache");
        response.setHeader("Expires", "0");
        
        processRequest(request, response);
    }
	
	protected void doPost(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		// TODO Auto-generated method stub
		String Message = "";
		String path = "C:\\DEV\\UploadedFile\\";
		
        
        
		try {
			Part part = request.getPart("file");
			String contentDisposition = part.getHeader("content-disposition");
			
			System.out.println("Disposition: " + contentDisposition);
			
			response.getWriter().println("test file Upload");
			Message = "Your File Was Uploaded Successfully " + GetFileName(part);
			part.write(GetFileName(part));
			path =  path + GetFileName(part);
		}catch(Exception ex) {
			Message = "Error Uploading The File " + ex.getMessage();
		}
		if(runOnce) {
			FileInputStream fileContent = new FileInputStream(path);
			ExcelColumnFilter filterExcel = new ExcelColumnFilter();
	        DateComparator compareDates = new DateComparator();
	        
	        compareDates.PrintCurrentDate(fileContent);
	        fileContent.close();
	        FileInputStream fileContent1 = new FileInputStream(path);
	        filterExcel.filterExcel(fileContent1);
	      
	        compareDates.getDateDifference();
	        fileContent1.close();
	        LocalDate currentDate = LocalDate.now();
	        FileInputStream fileContent2 = new FileInputStream("C:\\DEV\\UploadedFile\\"+currentDate+".xlsx");
	        DataGrouping SummaryTable = new DataGrouping();
	       
	        SummaryTable.createSummary(fileContent2, "C:\\DEV\\UploadedFile\\" + currentDate + ".xlsx");
	        
	        System.out.println("Summary Table Has been Created in Sheet - Summary Report");
	        fileContent2.close();
			/*
			 * request.setAttribute("message", Message);
			 * request.getRequestDispatcher("message.jsp").forward(request, response);
			 */
	        String filePath = "C:\\Users\\ADMIN\\eclipse-workspace\\DefectAgeing\\src\\main\\webapp\\message.html";
	        ConvertToHtml("C:\\DEV\\UploadedFile\\" + currentDate + ".xlsx", filePath);
	       

	        request.setAttribute("UploadedFile", Message);
			processRequest(request, response);
	        runOnce = false;
		}
		else {
			request.setAttribute("UploadedFile", Message);
			
			processRequest(request, response);
		}
		
	}
	private void processRequest(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
        // Set cache control headers
        response.setHeader("Cache-Control", "no-cache, no-store, must-revalidate");
        response.setHeader("Pragma", "no-cache");
        response.setHeader("Expires", "0");
        
        request.getRequestDispatcher("message.html").forward(request, response);
    }

	void ConvertToHtml (String excelFilePath,String htmlFilePath)  {
           try (Workbook workbook = new XSSFWorkbook(new FileInputStream(excelFilePath));
             PrintWriter writer = new PrintWriter(htmlFilePath)) {

            Sheet sheet = workbook.getSheet("Summary Report");

            Document doc = Jsoup.parse("<html><head></head><body></body></html>");
            Element body = doc.body();
            Element head = doc.head();

            Element table = doc.createElement("table");
            table.attr("style", "border-collapse: collapse;");
            body.appendChild(table);
            
            
            Element Former = doc.createElement("form");
            Former.attr("action", "UploadFile");
            Former.attr("method", "get");
            Former.attr("enctype", "multipart/form-data");
            body.appendChild(Former);
            
            Element h1Tag = doc.createElement("h1");
            Element ButtonReloader = doc.createElement("button");
            ButtonReloader.appendChild(h1Tag);
            h1Tag.appendText("ReloadThis");
            Former.appendChild(ButtonReloader);
            
            
            int NumCols = 0;
            int RemainingCols = 0;
            // Iterate over the rows and cells of the sheet
            for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
            	Row row = sheet.getRow(i);
            	
                Element tableRow = doc.createElement("tr");
                table.appendChild(tableRow);
                for (int j = 0; j < row.getPhysicalNumberOfCells(); j++) {
                	System.out.println(NumCols + "  " + RemainingCols);
                	Cell cell = row.getCell(j);
                	if(i == 0) {
                		NumCols++;
                	}
                	if(i != 0) {
                		RemainingCols--;
                	}
                    Element tableCell = doc.createElement("td");
                    tableCell.attr("style", "border: 1px solid black; padding: 5px;");
                    if(cell == null) {
                    	tableCell.text("nn");
                        tableRow.appendChild(tableCell);
                    	continue;
                    }
                    tableCell.text(cell.toString());
                    tableRow.appendChild(tableCell);
                }
                for(int k = RemainingCols; k > 0; k--) {
                	System.out.println(k);
                	Cell cell = row.getCell(NumCols - k);
                	Element tableCell = doc.createElement("td");
                    tableCell.attr("style", "border: 1px solid black; padding: 5px;");
                    if(cell == null) {
                    	tableCell.text("");
                        tableRow.appendChild(tableCell);
                    	continue;
                    }
                    tableCell.text(cell.toString());
                    tableRow.appendChild(tableCell);
                }
            	RemainingCols = NumCols;
            }
            Element script = doc.createElement("script");
            script.attr("src", "./UploadedFileProcess.js");
            doc.body().appendChild(script);
            doc.outputSettings().escapeMode(EscapeMode.xhtml);
            writer.write(doc.outerHtml());
            
            
            System.out.println("Excel converted to HTML successfully.");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

	 
	private String GetFileName(Part part) {
		String contentDisposition = part.getHeader("content-disposition");
		
		if(!contentDisposition.contains("filename=")) {
			return null;
		}
		
		int beginIndex = contentDisposition.indexOf("filename=") + 10;
		int endIndex = contentDisposition.length() - 1;
		
		return contentDisposition.substring(beginIndex, endIndex);
	}

}
