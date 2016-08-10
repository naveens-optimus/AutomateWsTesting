import com.eviware.soapui.impl.wsdl.WsdlInterface;
import com.eviware.soapui.impl.wsdl.WsdlOperation;
import com.eviware.soapui.impl.wsdl.WsdlProject;
import com.eviware.soapui.impl.wsdl.WsdlRequest;
import com.eviware.soapui.impl.wsdl.WsdlSubmit;
import com.eviware.soapui.impl.wsdl.WsdlSubmitContext;
import com.eviware.soapui.impl.wsdl.support.wsdl.WsdlImporter;
import com.eviware.soapui.model.iface.Operation;
import com.eviware.soapui.model.iface.Request.SubmitException;
import com.eviware.soapui.model.iface.Response;
import com.eviware.soapui.support.SoapUIException;

import java.io.*;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.*;
import org.apache.xerces.impl.xpath.XPath;
import org.apache.xmlbeans.XmlException;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Scanner;

import javax.activity.InvalidActivityException;
import javax.xml.xpath.XPathFactory;

import junit.textui.TestRunner;

import com.eviware.soapui.support.XmlHolder;


public class TestWS {

	public static void main(String[] args) throws IOException {
		
		System.out.println("Welcome user ! this suit will help you to automate webservice testing!!");
		System.out.println("Currently supported SOAP service testing!!");
		
		System.out.println("Please enter the excel file URL");
		System.out.println("");
		
		FileInputStream fsIP = null;
		Scanner reader = null;
		String excelFilePath = null;
		try{
			
			//read excel file path
			reader = new Scanner(System.in);  
			excelFilePath = reader.nextLine(); // Scans the next token of the input as an int.
			
			if(excelFilePath != null && !excelFilePath.equals("")){
				fsIP = new FileInputStream(new File(excelFilePath)); //Read the spreadsheet that needs to be updated
			}
			else{
				throw new InvalidActivityException("excel File path can not be blank");
			}
		}catch(IOException ex){
			
			System.err.println("Could not connect with the file, error: " + ex.getMessage());
			System.err.println("Fix the errors and start again.");
		}
		
		finally{
			reader.close();
			reader = null;
		}
		
		//Connect the excel file
		
		//Get the workbook instance for XLS file 
		HSSFWorkbook workbook = new HSSFWorkbook(fsIP);
	
		//Get first sheet from the workbook
		HSSFSheet sheet = workbook.getSheetAt(0);
		//read file to find a column with value WSDL
		
		Cell wsdlCell = findNextCellByCellData( sheet, "endpoint", 1);
		
		//pass the wsdl to the SoapUI api
		WsdlProject project = null;
		try {
			project = new WsdlProject();
		} catch (XmlException | SoapUIException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
        WsdlInterface[] wsdls = null;
		try {
			wsdls = WsdlImporter.importWsdl(project, wsdlCell.getStringCellValue() + "?wsdl");
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
        WsdlInterface wsdl = wsdls[0];
        

      	//Get active operations from the excel sheet
      	List<String> activeOperations = getOperationList(sheet);
        
      //iterate over all possible Soap requests in the WSDL
        for (Operation operation : wsdl.getOperationList()) {
        	
        	//Get operation Name
        	String operationName = operation.getName();
        	//Check if the operation is in the active operation list for the test automation
            if(!activeOperations.contains(operationName.toLowerCase())){
            	//No
            	//move to next operation
            	continue;
            }
        	
            
        	WsdlOperation op = (WsdlOperation) operation;
            
            //add a new Request
            WsdlRequest request = op.addNewRequest("Request");
            
            //Get Operation Name
            System.out.println("Executing request for OP:"+ operationName);
            System.out.println(op.createRequest(true));
            
            
            //Get SOap request
            XmlHolder xmlHolder = null;
            try {
            	xmlHolder = new XmlHolder(op.createRequest(true));
			} catch (XmlException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			};
            
			//Get user provided Tag details for this request from the excel
			Row opReqStartRow = findRowByCellData(sheet, operationName.toLowerCase().trim() + "-Request");
			
			//Get Test case tags row from Excel for the operation
			Row tagHeadingRow = sheet.getRow(opReqStartRow.getRowNum() + 1);
	        
			//Get user provided Tag details for this request from the excel
			Row opResponseStartRow = findRowByCellData(sheet, operationName.toLowerCase().trim() + "-Response");
			
	        Row resultTagHeadingRow = sheet.getRow(opResponseStartRow.getRowNum() + 1);
	        
	      	int testCaseRowIndex  = tagHeadingRow.getRowNum();
	      	
	      	int testResultRowIndex = resultTagHeadingRow.getRowNum();
	      	
	      	//iterate over number of test cases
	      	//Get number of test cases
	      	int noOfTestCases = findRowByCellData( sheet,  operationName.toLowerCase().trim() + "-EndTestCase").getRowNum() - testCaseRowIndex;
	      	
	      	//Which test case has been fired, to be used to update the related TestResult row
	      	int testResultRow = 0;
	      	
	      	//To be used for checking if the test case has to be executed
	      	Cell runCell = findColumnByCellData( tagHeadingRow, "Run", 0);
	      	
	      	//TODO: Make it dynamic so that the loop should work on test case length
	      	for(int i=0; i<noOfTestCases-1 ; i++){
	      		
	      	// move to next row
	            //TODO: Move it to Try/catch
	            testCaseRowIndex += 1;
	            Row testCaseRow = sheet.getRow(testCaseRowIndex);
	            
	          //Check if the test case has to be executed
	    		
	    		String isTestCaseToBeExecuted = testCaseRow.getCell(runCell.getColumnIndex()).getStringCellValue();
	    		
	    		//if Yes, fill the possible tag values
	    		if(isTestCaseToBeExecuted.equalsIgnoreCase("y")){
	    			
	    			 //Set next test result row
	    			testResultRow = testCaseRowIndex - tagHeadingRow.getRowNum();
	    			
	    			XmlHolder filledRequest = fillRequest(xmlHolder, sheet, testCaseRow, tagHeadingRow);
	                
	                System.out.println("Filled request= " + filledRequest.getXml());
	                
	                //replace WSDL request and execute it
	                if(filledRequest != null){
	    	            request.setRequestContent(filledRequest.getXml().replace("?", ""));
	    	            
	    	            //execute
	    	         // submit the request
	    	            WsdlSubmit submit = null;
	    				try {
	    					submit = (WsdlSubmit) request.submit( new WsdlSubmitContext(op), false );
	    				} catch (SubmitException e) {
	    					// TODO Auto-generated catch block
	    					e.printStackTrace();
	    				}
	    	           
	    	            // wait for the response
	    	            Response response = submit != null ? submit.getResponse() : null;
	    	           
	    	            //  print the response
	    	            String content = response != null ? response.getContentAsString() : null;
	    	            
	    	            XmlHolder responseXmlHolder = null;
	    				try {
	    					responseXmlHolder = new XmlHolder(content);
	    				} catch (XmlException e) {
	    					// TODO Auto-generated catch block
	    					e.printStackTrace();
	    				}
	    	            
	    				//Save response in excel
	    	            //Save response to next row
	    	            writeResponse(responseXmlHolder, resultTagHeadingRow, sheet.getRow(resultTagHeadingRow.getRowNum() + testResultRow));
	    	            
	    	            System.out.println("Request Response= \n " + content );
	                }
	    		}
	      	}
	      	
            
            //System.out.println("Response:");
            //System.out.println(op.createResponse(true));
        }
        
        
      	
      	//Save file and close
      	fsIP.close(); //Close the InputStream
        
        FileOutputStream output_file = new FileOutputStream(new File(excelFilePath));  //Open FileOutputStream to write updates
          
        workbook.write(output_file); //write changes
          
        output_file.close(); 
        
        //finished the testing
        System.out.println("All the test cases executed and results are written in the excel file, please read the file to get more info.");
        System.out.println("System is shutting down now.................................................................................................................");
        System.out.println("Happy test automation");
	}
	
	
	//Method to Fill data in the 
	private static XmlHolder fillRequest(XmlHolder request, HSSFSheet workSheet, Row testCaseValueRow, Row testCaseTagNameRow){
				
		System.out.println("filling test data from row= " + testCaseValueRow.getRowNum()+1);
		
		int testCaseCounter = testCaseValueRow.getRowNum() - testCaseTagNameRow.getRowNum();
		
		//For tag and value row, iterate through each columns
		Iterator<Cell> tagCellIterator = testCaseTagNameRow.cellIterator();
		
		while(tagCellIterator.hasNext() ) {
			
			Cell tagCell = tagCellIterator.next();
			
			//Check if encounter end of Tag Name
			if(tagCell.getStringCellValue().equalsIgnoreCase("run")){
				//break the loop
				break;
			}
			
			//cellCounter++;
			if(tagCell.getColumnIndex() >= 2){
				System.out.println("filling tagValue " + tagCell.getStringCellValue() + "\t\t" );
				
				//Get value cell for the tag
					Cell valCell = testCaseValueRow.getCell(tagCell.getColumnIndex());
					
					
					//fill tag value
					try {
						switch( valCell.getCellType() ) {
							case Cell.CELL_TYPE_BOOLEAN:
								//System.out.println(tagCell.getBooleanCellValue() + "\t\t");
								request.setNodeValue("//*:" + tagCell.getStringCellValue() , valCell.getBooleanCellValue());
								break;
							case Cell.CELL_TYPE_NUMERIC:
								//System.out.println(tagCell.getNumericCellValue() + "\t\t");
							
								request.setNodeValue("//*:" + tagCell.getStringCellValue() , valCell.getNumericCellValue());
							
								break;
							case Cell.CELL_TYPE_STRING:
								System.out.println(tagCell.getStringCellValue() + "\t\t");
								
								//XPath xpath = new XPathFactory.newInstance().newXPath();
								//System.out.println("FromCurrency= " + request.getNodeValue("//*:" + tagCell.getStringCellValue()));
								request.setNodeValue("//*:" + tagCell.getStringCellValue() , valCell.getStringCellValue());
								break;
						}
						
						System.out.println("xmlHolder= " + request.getXml());
					} catch (XmlException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
				//}
			}
		}
		System.out.println("");
		
		//return filled request
		return request;
		
	}
	
	
	private String readExcel(String fileAddress, String columnToRead){
		String returnVal = "";
		
        System.out.println("file processing start");
		try {
			
			 FileInputStream fsIP= new FileInputStream(new File("D:\\output1.xls")); //Read the spreadsheet that needs to be updated
			
			//Get the workbook instance for XLS file 
			HSSFWorkbook workbook = new HSSFWorkbook(fsIP);
		
			//Get first sheet from the workbook
			HSSFSheet sheet = workbook.getSheetAt(0);
		
			System.out.println("sheet.getLastRowNum()= "  + sheet.getLastRowNum());
			
			int rowCounter = 0;
			int cellCounter = 3;
			//Iterate through each rows from first sheet
			Iterator<Row> rowIterator = sheet.iterator();
			while(rowIterator.hasNext()) {
				Row row = rowIterator.next();
				//rowCounter++;
				System.out.println("printing row= " + rowCounter);
				//For each row, iterate through each columns
				Iterator<Cell> cellIterator = row.cellIterator();
				while(cellIterator.hasNext()) {
					
					Cell cell = cellIterator.next();
					//cellCounter++;
					System.out.println("printing cell= " + cellCounter);
					
					switch(cell.getCellType()) {
						case Cell.CELL_TYPE_BOOLEAN:
							System.out.println(cell.getBooleanCellValue() + "\t\t");
							break;
						case Cell.CELL_TYPE_NUMERIC:
							System.out.println(cell.getNumericCellValue() + "\t\t");
							break;
						case Cell.CELL_TYPE_STRING:
							System.out.println(cell.getStringCellValue() + "\t\t");
							break;
					}
				}
				System.out.println("");
			}
		
			//Write to file
			Cell cell = null; // declare a Cell object
		
			System.out.println("rowCount= " + rowCounter);
		    System.out.println("cellCount= " + cellCounter);
		        //cell = sheet.getRow(sheet.getLastRowNum()).getCell(cellCounter-1);   // Access the second cell in second row to update the value
			cell = sheet.getRow(1).getCell(cellCounter-1); 
		
			Object cellValue = "";
		        switch(cell.getCellType()) {
						case Cell.CELL_TYPE_BOOLEAN:
							cellValue= cell.getBooleanCellValue() ;
							break;
						case Cell.CELL_TYPE_NUMERIC:
							cellValue= cell.getNumericCellValue();
							break;
						case Cell.CELL_TYPE_STRING:
							cellValue= cell.getStringCellValue();
							break;
					}
		        if(cellValue == null || cellValue == "")
		        	cell.setCellValue("OverRide Last Name");  // Get current cell value value and overwrite the value
		        else{
		        		//Set cell to next column cell
		        		cell = sheet.getRow(sheet.getLastRowNum()).getCell(cellCounter); 
		
		        		//Check if cell is null?
		        		if(cell == null){
		        			//Create Cell
		        			sheet.getRow(sheet.getLastRowNum()).createCell(cellCounter);
		        			cell = sheet.getRow(sheet.getLastRowNum()).getCell(cellCounter);
		        			}
		        		cell.setCellValue("OverRide Last Name"); 
		        	}
		          
		        fsIP.close(); //Close the InputStream
		         
		        FileOutputStream output_file =new FileOutputStream(new File("D:\\output1.xls"));  //Open FileOutputStream to write updates
		          
		        //workbook.write(output_file); //write changes
		          
		        output_file.close();  //close the stream  
		       
			
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		
		
		System.out.println("file processing complete");
		
		return returnVal;
	}
	
	
	private void writeExcel(String fileAddress, Integer row, Integer col, String value){
		try{
			
		}
		catch(Exception ex){
			
		}
		finally{
			
		}
	}
	
	private static Row findRowByCellData(HSSFSheet sheet, String cellContent) {
	    for (Row row : sheet) {
	        for (Cell cell : row) {
	            if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
	                if (cell.getRichStringCellValue().getString().trim().toLowerCase().equals(cellContent.toLowerCase())) {
	                    return row;
	                }
	            }
	        }
	    }               
	    return null;
	}
	
	private static int findColumnByCellData(HSSFSheet sheet, String cellContent) {
	    for (Row row : sheet) {
	        for (Cell cell : row) {
	            if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
	                if (cell.getRichStringCellValue().getString().trim().equals(cellContent)) {
	                    return cell.getColumnIndex();
	                }
	            }
	        }
	    }               
	    return 0;
	}
	
	/**
	 * Gets a cell from the HSSFSheet using the cellContent
	 * @param sheet
	 * @param cellContent
	 * @param afterNoOfPlaces - Gets cell after or before the provided position, 0 = same cell with the cell content
	 * @return Cell
	 */
	private static Cell findColumnByCellData(HSSFSheet sheet, String cellContent, int afterNoOfPlaces) {
	    for (Row row : sheet) {
	        for (Cell cell : row) {
	            if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
	                if (cell.getRichStringCellValue().getString().trim().equals(cellContent)) {
	                    return cell.getRow().getCell( cell.getColumnIndex() + afterNoOfPlaces );
	                }
	            }
	        }
	    }               
	    return null;
	}
	
	/**
	 * Gets a cell from the Row using the cellContent
	 * @param sheet
	 * @param cellContent
	 * @param afterNoOfPlaces - Gets cell after or before the provided position, 0 = same cell with the cell content
	 * @return Cell
	 */
	private static Cell findColumnByCellData(Row row, String cellContent, int afterNoOfPlaces) {
	   for (Cell cell : row) {
	            if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
	                if (cell.getRichStringCellValue().getString().toLowerCase().trim().equals(cellContent.toLowerCase())) {
	                    return cell.getRow().getCell( cell.getColumnIndex() + afterNoOfPlaces );
	                }
	            }
	        }               
	    return null;
	}
	
	/**
	 * 
	 * @param sheet
	 * @param cellContent
	 * @param afterPosition
	 * @return
	 */
	private static Cell findNextCellByCellData(HSSFSheet sheet, String cellContent, int afterPosition) {
	    for (Row row : sheet) {
	        for (Cell cell : row) {
	            if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
	                if (cell.getRichStringCellValue().getString().trim().equals(cellContent)) {
	                    return row.getCell(cell.getColumnIndex() + afterPosition) ;
	                }
	            }
	        }
	    }               
	    return null;
	}
	
	/**
	 * 
	 * @param sheet
	 * @param cellContent
	 * @param downRows
	 * @return
	 */
	private static Cell findDownCellByCellData(HSSFSheet sheet, String cellContent, int downRows) {
	    for (Row row : sheet) {
	        for (Cell cell : row) {
	            if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
	                if (cell.getRichStringCellValue().getString().trim().equals(cellContent)) {
	                    return sheet.getRow(row.getRowNum() + downRows).getCell(cell.getColumnIndex()) ;
	                }
	            }
	        }
	    }               
	    return null;
	}

	/***
	 * 
	 * @param response
	 * @param responseTagRow
	 * @param responseRow
	 */
	private static void writeResponse(XmlHolder response, Row responseTagRow, Row responseRow){
		
		//Get response value for Response tags 
		System.out.println("filling test data from row= " + responseRow.getRowNum()+1);
		
		
		//int testCaseCounter = testCaseValueRow.getRowNum() - testCaseTagNameRow.getRowNum();
		
		if(response == null || response.getXml().length() > 1){
			System.out.println("response null, check the request");
		}
		
		//For tag and value row, iterate through each columns
		Iterator<Cell> tagCellIterator = responseTagRow.cellIterator();
		
		while(tagCellIterator.hasNext() ) {
			
			Cell tagCell = tagCellIterator.next();
			
			
			//Check if encounter end of Tag Name
			if(tagCell.getStringCellValue().equalsIgnoreCase("EndResult")){
				//break the loop
				break;
			}
			
			
			
			//cellCounter++;
			if(tagCell.getColumnIndex() >= 2){
				
				if(tagCell.getStringCellValue().equals("")){
					continue;
				}
				
				System.out.println("filling response tagValue " + tagCell.getStringCellValue() + "\t\t" );
				
				//Get value cell for the tag
					Cell valCell = responseRow.getCell(tagCell.getColumnIndex());
					
					
					//fill tag value
					try {
						
						//Get value from the XmlHolder of operation Response
						System.out.println("Response xmlHolder= " + response.getXml());
						
						valCell.setCellValue(response.getNodeValue("//*:" + tagCell.getStringCellValue()));
						
					} catch (XmlException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
				//}
			}
		}
		System.out.println("Response write done");
		
	}
	
	
	/**
	 * 
	 * @param sheet
	 * @return
	 */
	private static List<String> getOperationList(HSSFSheet sheet){
		List<String> operationList = new ArrayList<String>();
		
		//Get operation list
		//findout the cell below the cell with value OperationName
		Cell startCell = findDownCellByCellData( sheet, "OperationName", 1);
		
		//get the start index of the iterator
		int startIndex = startCell.getRowIndex();
		
		//get the last index
		Cell endCell = findColumnByCellData(sheet, "EndOperation", 0);
		
		int endIndex = endCell.getRowIndex();
		
		//iterate on the range of operations
		for(; startIndex < endIndex ; startIndex++){
			
			//Check if operation is active
			//if yes, add it in the list 
			if(startCell.getRow().getCell(startCell.getColumnIndex() + 1).getStringCellValue().equalsIgnoreCase("Y")){
				operationList.add(startCell.getStringCellValue().trim().toLowerCase());
				
			}
			
			//move to down cell
			startCell = sheet.getRow(startCell.getRowIndex()+1).getCell(startCell.getColumnIndex());
		}
		return operationList;
	}
}

