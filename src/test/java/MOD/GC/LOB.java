package MOD.GC;

	import java.io.File;
	import java.io.FileInputStream;
	import java.io.FileOutputStream;
	import java.io.IOException;
	import java.io.InputStream;
	import java.math.BigDecimal;
	import java.util.ArrayList;
	import java.util.Iterator;
	import java.util.List;

	import org.apache.poi.hssf.usermodel.HSSFWorkbook;
	import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
	import org.apache.poi.ss.usermodel.Cell;
	import org.apache.poi.ss.usermodel.Row;
	import org.apache.poi.ss.usermodel.Sheet;
	import org.apache.poi.ss.usermodel.Workbook;
	import org.apache.poi.xssf.usermodel.XSSFCell;
	import org.apache.poi.xssf.usermodel.XSSFRow;
	import org.apache.poi.xssf.usermodel.XSSFSheet;
	import org.apache.poi.xssf.usermodel.XSSFWorkbook;
	import org.apache.poi.xssf.usermodel.XSSFWorkbookType;
	import org.testng.annotations.Test;

	import com.monitorjbl.xlsx.StreamingReader;

	public class LOB {
		@Test
	     public static void main() throws IOException, InvalidFormatException{

	         

			
			InputStream is = new FileInputStream(new File("D:\\backup\\GC\\WVFLORIDA2.xlsx"));
				
			Workbook workbook = StreamingReader.builder()
			        .rowCacheSize(100)    // number of rows to keep in memory (defaults to 10)
			        .bufferSize(4096)     // buffer size to use when reading InputStream to file (defaults to 1024)
			        .open(is);            // InputStream or File for XLSX file (required)
			Sheet sheet = workbook.getSheet("LOB Item Mapping");
			
			

	      	int rowNum = 10000;
	        //int rowNum = sheet.getLastRowNum() + 1;
	        //int colNum = sheet.getRow(0).getLastCellNum();
	        int colNum = 26;
	        int AmdocsCorpIDHeaderIndex = -1, AmdocsRateCodeHeaderIndex = -1;
	        int IsInternetLOBCodeHeaderIndex = -1, IsPhoneLOBHeaderIndex = -1;
	        int CSGCodeHeaderIndex = -1, DoNotConvertHeaderIndex = -1;
	        int ProvisionedIndex = -1, RateHeaderIndex = -1;
	        
	        
	        
	        
			for (Row r : sheet) {
				  for (Cell c : r) {
					  
					  for (int j = 0; j < colNum; j++) {
			            	Cell cell = r.getCell(j);
			                String cellValue = cellToString(cell);
			                if ("Amdocs Corp ID".equalsIgnoreCase(cellValue)) {
			                	AmdocsCorpIDHeaderIndex = j;
			                } else if ("Amdocs Rate Code".equalsIgnoreCase(cellValue)) {
			                	AmdocsRateCodeHeaderIndex = j;
			                }else if ("LOB I (Internet)".equalsIgnoreCase(cellValue)) {
			                	IsInternetLOBCodeHeaderIndex = j;
			                }else if ("LOB T (Phone)".equalsIgnoreCase(cellValue)) {
			                	IsPhoneLOBHeaderIndex = j;
			                }else if ("CSG Code(s)".equalsIgnoreCase(cellValue)) {
			                	CSGCodeHeaderIndex = j;
			                }else if ("Do Not Convert (Y/N)".equalsIgnoreCase(cellValue)) {
			                	DoNotConvertHeaderIndex = j;
			                }else if ("Provisioned (Y/N)".equalsIgnoreCase(cellValue)) {
			                	ProvisionedIndex = j;
			                }else if ("Rate(s) for Standard to Bulk/NS Mapping(s)".equalsIgnoreCase(cellValue)) {
			                	RateHeaderIndex = j;
			                    
			               
			            }
			               
			            }
					 
					  if (AmdocsCorpIDHeaderIndex == -1 || AmdocsRateCodeHeaderIndex == -1) {
			                System.out.println("LOLA");
			            }
					  break;
					  
				  }
				  break;
				}     
	        
		
	    
	                     
	         // createnew workbook
			
			
	            XSSFWorkbook workbook1 = new XSSFWorkbook();
	            // Create a blank sheet
	            XSSFSheet sheet1 = workbook1.createSheet("LOB Item Mapping");
	            XSSFRow newRow = sheet1.createRow(0); 
	            int HeadIndex = 0;
	            newRow.createCell(HeadIndex++).setCellValue("AmdocsCorpID");
	            newRow.createCell(HeadIndex++).setCellValue("AmdocsRateCode");
	            newRow.createCell(HeadIndex++).setCellValue("IsInternetLOB");
	            newRow.createCell(HeadIndex++).setCellValue("IsPhoneLOB");
	            newRow.createCell(HeadIndex++).setCellValue("CSGCode");
	            newRow.createCell(HeadIndex++).setCellValue("DoNotConvert");
	            newRow.createCell(HeadIndex++).setCellValue("Provisioned");
	            newRow.createCell(HeadIndex++).setCellValue("Rate");
	            
	            String IsInternetLOB = null;
	            String IsPhoneLOB  = null;
	            String CSGCode = null;
	            String DoNotConvert = null;
	            String Provisioned = null;
	            String Rate = null;
	            
	            
	            
	            int i = 1;
	            
	            for (Row r : sheet) {
	  		// for (Cell c : r) {
	           
	                
	            	
	                String AmdocsCorpID = cellToString(r.getCell(AmdocsCorpIDHeaderIndex));
	                String AmdocsRateCode = cellToString(r.getCell(AmdocsRateCodeHeaderIndex));
	                AmdocsRateCode =  AmdocsRateCode.replaceAll("\n","|");
	                
	                
	                if(r.getCell(IsInternetLOBCodeHeaderIndex )== null)
	                {	
	                }
	                else               	
	                IsInternetLOB = cellToString(r.getCell(IsInternetLOBCodeHeaderIndex ));
	                
	                
	                if(r.getCell(IsPhoneLOBHeaderIndex )== null)
	                {
	                }
	                else
	                	IsPhoneLOB = cellToString(r.getCell(IsPhoneLOBHeaderIndex ));
	                
	                if(r.getCell(CSGCodeHeaderIndex )== null)
	                {
	                }
	                else
	                	CSGCode = cellToString(r.getCell(CSGCodeHeaderIndex));
	                
	               
	                if(r.getCell(DoNotConvertHeaderIndex )== null)
	                {
	                }
	                else
	                	DoNotConvert = cellToString(r.getCell(DoNotConvertHeaderIndex ));
	                
	                if(r.getCell(ProvisionedIndex  )== null)
	                {
	                }
	                else
	                	Provisioned = cellToString(r.getCell(ProvisionedIndex));
	                
	                if(r.getCell(RateHeaderIndex)== null)
	                {
	                }
	                else
	                	Rate = cellToString(r.getCell(RateHeaderIndex));
	                
	                
	                int cellIndex = 0;
	                
	                if(AmdocsCorpID.contains("1624")||AmdocsCorpID.contains("ALL"))
	                {
	                //Create a newRow object for the output excel. 
	                //We begin for i = 1, because of the headers from the input excel, so we go minus 1 in the new (no headers).
	                //If for the output we need headers, add them outside this for loop, and go with i, not i-1
	                
	                newRow = sheet1.createRow(i);  
	                
	                if(AmdocsCorpID.contains("ALL"))
	                {
	                newRow.createCell(cellIndex++).setCellValue("1624");	                	
	                }
	                else
	                newRow.createCell(cellIndex++).setCellValue(AmdocsCorpID);
	                newRow.createCell(cellIndex++).setCellValue(AmdocsRateCode);
	                newRow.createCell(cellIndex++).setCellValue(IsInternetLOB);
	                newRow.createCell(cellIndex++).setCellValue(IsPhoneLOB);
	                newRow.createCell(cellIndex++).setCellValue(CSGCode);
	                newRow.createCell(cellIndex++).setCellValue(DoNotConvert);
	                newRow.createCell(cellIndex++).setCellValue(Provisioned);
	                newRow.createCell(cellIndex++).setCellValue(Rate);
	                i++;
	                }
	           // }
	  			  
	            }
	            
	            FileOutputStream fos = new FileOutputStream(new File("D:\\backup\\GC\\GC_LOB1.xlsx"));
	            workbook1.write(fos);
	            fos.close();
	            
	            }
		
		
	            public static String cellToString(Cell cell) 
	            {
	                int type;
	                Object result = null;
	                type = cell.getCellType();

	                switch (type) {

	                case XSSFCell.CELL_TYPE_NUMERIC:
	                    result = BigDecimal.valueOf(cell.getNumericCellValue()).intValue();

	                    break;
	                case XSSFCell.CELL_TYPE_STRING:
	                    result = cell.getStringCellValue();
	                    break;
	                case XSSFCell.CELL_TYPE_BLANK:
	                    result = "";
	                    break;
	                case XSSFCell.CELL_TYPE_FORMULA:
	                    result = cell.getCellFormula();
	                }
	                	System.out.println(result);
	                return result.toString();
	            }}

	        


	   	

