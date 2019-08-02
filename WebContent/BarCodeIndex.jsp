<%@page import="java.sql.ResultSet"%>
<%@page import="java.sql.Statement"%>
<%@page import="java.sql.*"%>
<%@page import="java.io.*"%>
<%@page import="java.util.*"%>
<%@page import="org.apache.poi.hssf.usermodel.*"%>
<%@page import="org.apache.poi.ss.usermodel.Cell"%>;
<%@page import="org.apache.poi.ss.usermodel.*"%>;
<%@page import="org.apache.poi.xssf.usermodel.*"%>;    
<%@ page import="org.apache.commons.fileupload.*" %>
    <%@ page import="org.apache.commons.fileupload.disk.*" %>
    <%@ page import="org.apache.commons.fileupload.servlet.*" %>
    <%@ page import="org.apache.commons.io.output.*" %>
<%@ page language="java" contentType="text/html;charset=ISO-8859-1"
 pageEncoding="ISO-8859-1"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<title>BarCodeGenerator</title>
</head>
<body>
<center>
<table cellpadding="10" border="2" width=80%>
 <%
       File file =new File("");
       int maxFileSize = 5000 * 1024;
       int maxMemSize = 5000 * 1024;

       String fileName="";
       String filePath = "C:/JavaPrograms/BarCodeGenerator/Files/";
     
       String contentType = request.getContentType();
       if ((contentType.indexOf("multipart/form-data") >= 0)) {
          DiskFileItemFactory factory = new DiskFileItemFactory();
          factory.setSizeThreshold(maxMemSize);
          factory.setRepository(new File("C:/JavaPrograms/BarCodeGenerator/Files"));
          ServletFileUpload upload = new ServletFileUpload(factory);
          upload.setSizeMax( maxFileSize );
          try{ 
             List fileItems = upload.parseRequest(request);
             Iterator i = fileItems.iterator();
             while ( i.hasNext () ) 
             {
                FileItem fi = (FileItem)i.next();
                if ( !fi.isFormField () )  {
                    String fieldName = fi.getFieldName();
                    fileName = fi.getName();
                    boolean isInMemory = fi.isInMemory();
                    long sizeInBytes = fi.getSize();
                    file = new File( filePath + fileName) ;
                    fi.write( file ) ;
                    //out.println("Uploaded Filename: " + filePath + fileName + "<br>");
                }
             }
          }catch(Exception ex) {
            out.println(ex);
          }
       }else{
          out.println("<p>No file uploaded</p>"); 
       }
    

	String sheetName="Sheet1";
	
	//Create an object of FileInputStream class to read excel file
	
	FileInputStream inputStream = new FileInputStream(file);
	
	Workbook barCodeWorkbook = null;
	//Find the file extension by splitting file name in substring  and getting only extension name
	
	String fileExtensionName = fileName.substring(fileName.indexOf("."));
	barCodeWorkbook = new XSSFWorkbook(inputStream);
	
		Sheet barCodeSheet = barCodeWorkbook.getSheet(sheetName);
		//Find number of rows in excel file
		int rowCount = barCodeSheet.getLastRowNum()-barCodeSheet.getFirstRowNum();
		//Create a loop over all the rows of excel file to read it
		for (int i = 0; i < rowCount+1; i++) {
	    Row row = barCodeSheet.getRow(i);
	    boolean flag=false;
	    //Create a loop to print cell values in a row
			%> <tr><% 
			for (int j = 0; j < row.getLastCellNum(); j++) {
		    	switch (row.getCell(j).getCellType())
		    	{
		    		case Cell.CELL_TYPE_NUMERIC:        
		        		%>
		       			<td>
		        		<%=row.getCell(j).getNumericCellValue()%>
		        		</td><% 
			    		break;
			    	case Cell.CELL_TYPE_STRING:
			 	       if(j==row.getLastCellNum()-1 && i>0)
				    	{
				    		%><td>
				    		 <img alt="my Image" src="BarCode?user=<%=row.getCell(j).getStringCellValue()%>" height='60px' width='400px'><br><br>
				    		 </td><% 
				    	}
			 	       else
			 	       {
			 	    	  %>
				 	        <td>
				 	        <%=row.getCell(j).getStringCellValue()%>
				 	        </td><% 
			 	       }
			 	        
			    		break;
			    }
		    	
		    }
			%> </tr><% 
		}  
 %>
 </table>
 <button onclick="myFunction()" style="margin-top:10px;border-radius:5px;height:25px">Save or Print this Page</button>

<script>
function myFunction() {
  window.print();
}
</script>
 </html>