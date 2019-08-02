

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.PrintWriter;
import java.util.Iterator;

import javax.servlet.ServletException;
import javax.servlet.annotation.WebServlet;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.*;

/**
 * Servlet implementation class ExcelDataFetching
 */
@WebServlet("/ExcelDataFetching")
public class ExcelDataFetching extends HttpServlet {
	private static final long serialVersionUID = 1L;
       
    /**
     * @see HttpServlet#HttpServlet()
     */
    public ExcelDataFetching() {
        super();
        // TODO Auto-generated constructor stub
    }

	/**
	 * @see HttpServlet#doGet(HttpServletRequest request, HttpServletResponse response)
	 */
	protected void doGet(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		// TODO Auto-generated method stub
		
	      PrintWriter out=response.getWriter();
	 
	      try{
	    	  FileInputStream file;
	    	  file = new FileInputStream(new File("F:\\USNs.xlsx"));

	    	  //Create Workbook instance holding reference to .xlsx file
	    	  XSSFWorkbook workbook = new XSSFWorkbook(file);

	    	  //Get first/desired sheet from the workbook
	    	  XSSFSheet sheet = workbook.getSheetAt(0);

	    	  //Iterate through each rows one by one
	    	  Iterator<Row> rowIterator = sheet.iterator();
	    	  while (rowIterator.hasNext()){
	    	  Row row = rowIterator.next();
	    	  //For each row, iterate through all the columns
	    	  Iterator<Cell> cellIterator = row.cellIterator();

	    	  while (cellIterator.hasNext()){
	    	  Cell cell = cellIterator.next();

	    	  //Check the cell type and format accordingly
	    	  switch (cell.getCellType()){
	    	  case Cell.CELL_TYPE_NUMERIC:
	    	  out.print(cell.getNumericCellValue() + "t");
	    	  break;
	    	  case Cell.CELL_TYPE_STRING:
	    	  out.print(cell.getStringCellValue() + "t");
	    	  break;
	    	  }
	    	  }
	    	  out.println("");
	    	  }
	    	  file.close();
	    	  } 
	    	  catch (Exception e) 
	    	  {
	    	  e.printStackTrace();
	    	  }
	    	  
	
	}

	/**
	 * @see HttpServlet#doPost(HttpServletRequest request, HttpServletResponse response)
	 */
	protected void doPost(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		// TODO Auto-generated method stub
	}

}
