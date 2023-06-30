

import java.io.IOException;
import java.io.InputStream;
import java.time.LocalDate;
import java.util.List;

import javax.servlet.ServletException;
import javax.servlet.annotation.WebServlet;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.tomcat.util.http.fileupload.FileItem;
import org.apache.tomcat.util.http.fileupload.FileUploadException;
import org.apache.tomcat.util.http.fileupload.disk.DiskFileItemFactory;
import org.apache.tomcat.util.http.fileupload.servlet.ServletFileUpload;
import org.apache.tomcat.util.http.fileupload.servlet.ServletRequestContext;


/**
 * Servlet implementation class DefectAgeHandlerServlet
 */
@WebServlet(name = "DefectAgeHandlerServlet", urlPatterns = { "/DefectAge/upload" })
public class DefectAgeHandlerServlet extends HttpServlet {
	public static InputStream fileContent;
	private static final long serialVersionUID = 1L;
       
    /**
     * @see HttpServlet#HttpServlet()
     */
    public DefectAgeHandlerServlet() {
        super();
        // TODO Auto-generated constructor stub
    }

	/**
	 * @see HttpServlet#doGet(HttpServletRequest request, HttpServletResponse response)
	 */
    protected void doGet(HttpServletRequest request, HttpServletResponse response)
            throws ServletException, IOException {
        response.setContentType("text/html");

        // Check if the request contains a file upload
        boolean isMultipart = ServletFileUpload.isMultipartContent(request);

        if (isMultipart) {
            // Create a factory for disk-based file items
            DiskFileItemFactory factory = new DiskFileItemFactory();

            // Create a new file upload handler
            ServletFileUpload upload = new ServletFileUpload(factory);

            try {
                // Parse the request to get file items
                List<FileItem> fileItems = upload.parseRequest(new ServletRequestContext(request));

                for (FileItem fileItem : fileItems) {
                    if (!fileItem.isFormField()) {
                        // Process the uploaded file
                        InputStream fileContent = fileItem.getInputStream();

                        // Process the Excel file and generate the summary table
                        String summaryTable = processFile(fileContent);

                        // Set the summary table as a request attribute
                        request.setAttribute("summaryTable", summaryTable);

                        // Forward the request to the same HTML page
                        request.getRequestDispatcher("coreflex.html").forward(request, response);
                        return;
                    }
                }
            } catch (FileUploadException e) {
                e.printStackTrace();
            }
        }

        // If no file was uploaded or an error occurred, display an error message
        response.getWriter().println("Error: No file uploaded or an error occurred.");
    }

	/**
	 * @see HttpServlet#doPost(HttpServletRequest request, HttpServletResponse response)
	 */
	protected void doPost(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		 doGet(request, response);
		    }
		  
		

	private String processFile(InputStream fileContent) {
		LocalDate currentDate = LocalDate.now();
		ExcelColumnFilter filterExcel = new ExcelColumnFilter();
		DateComparator compareDates = new DateComparator();
		compareDates.PrintCurrentDate(fileContent);
		filterExcel.filterExcel(fileContent);
		compareDates.getDateDifference();
		System.out.println("Date Difference has been entered in file - DefectAgeing"+currentDate+".xlsx");
		DataGrouping summaryTable = new DataGrouping();
		summaryTable.createSummary();
		System.out.println("Summary Table Has been Created in Sheet - Summary Report");
		return DateComparator.filePath;
	}
}
