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
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
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

public class RES {
	@Test
     public static void main() throws IOException, InvalidFormatException{

         

		
		InputStream is = new FileInputStream(new File("D:\\backup\\GC\\WVFLORIDA2.xlsx"));
			
		Workbook workbook = StreamingReader.builder()
		        .rowCacheSize(100)    // number of rows to keep in memory (defaults to 10)
		        .bufferSize(4096)     // buffer size to use when reading InputStream to file (defaults to 1024)
		        .open(is);            // InputStream or File for XLSX file (required)
		Sheet sheet = workbook.getSheet("Res Item Mapping");
		
		

      	//int rowNum = sheet.getLastRowNum() + 1;
        //int colNum = sheet.getRow(0).getLastCellNum();
        int colNum = 20;
        int AmdocsCorpIDHeaderIndex = -1, AmdocsRateCodeHeaderIndex = -1;
        int CSGCodeHeaderIndex = -1, DoNotConvertHeaderIndex = -1;
        int CategoryHeaderIndex = -1, ProvisionedHeaderIndex = -1;
        
        
        
        
		for (Row r : sheet) {
			  for (Cell c : r) {
				  
				  for (int j = 0; j < colNum; j++) {
		            	Cell cell = r.getCell(j);
		                String cellValue = cellToString(cell);
		                if ("AmdocsCorpID".equalsIgnoreCase(cellValue)) {
		                	AmdocsCorpIDHeaderIndex = j;
		                } else if ("AmdocsRateCode".equalsIgnoreCase(cellValue)) {
		                	AmdocsRateCodeHeaderIndex = j;
		                }else if ("CSGCode(s)".equalsIgnoreCase(cellValue)) {
		                	CSGCodeHeaderIndex = j;
		                }else if ("DoNotConvert(Y/N)".equalsIgnoreCase(cellValue)) {
		                	DoNotConvertHeaderIndex = j;
		                }else if ("Category".equalsIgnoreCase(cellValue)) {
		                	CategoryHeaderIndex = j;
		                }else if ("Provisioned(Y/N)".equalsIgnoreCase(cellValue)) {
		                	ProvisionedHeaderIndex = j;
		                    
		               
		            }
		               
		            }
				 
				  if (AmdocsCorpIDHeaderIndex == -1 || AmdocsRateCodeHeaderIndex == -1) {
		                System.out.println("LOLA");
		            }
				  break;
				  
				  
			    //System.out.println(c.getStringCellValue());
			  }
			  break;
			}     
        
	
    
                     
         // create new workbook
		
		
            XSSFWorkbook workbook1 = new XSSFWorkbook();
            // Create a blank sheet
            XSSFSheet sheet1 = workbook1.createSheet("Res Item Mapping");
                    
            XSSFRow newRow = sheet1.createRow(0); 
            int HeadIndex = 0;
            newRow.createCell(HeadIndex++).setCellValue("AmdocsCorpID");
            newRow.createCell(HeadIndex++).setCellValue("AmdocsRateCode");
            newRow.createCell(HeadIndex++).setCellValue("Category");
            newRow.createCell(HeadIndex++).setCellValue("CSGCode");
            newRow.createCell(HeadIndex++).setCellValue("DoNotConvert");
            newRow.createCell(HeadIndex++).setCellValue("Provisioned");            
            String Category = null;
            String CSGCode  = null;
            String DoNotConvert = null;
            String Provisioned = null;
            
            
            
            int i = 1;
            
            for (Row r : sheet) {
  		// for (Cell c : r) {
           
                
            	
                String AmdocsCorpID = cellToString(r.getCell(AmdocsCorpIDHeaderIndex));
                
             
              
                String AmdocsRateCode = cellToString(r.getCell(AmdocsRateCodeHeaderIndex));                
                AmdocsRateCode =  AmdocsRateCode.replaceAll("\n","|");
                
                if(r.getCell(CategoryHeaderIndex)== null)
                {	
                }
                else               	
                Category = cellToString(r.getCell(CategoryHeaderIndex));
                
                
                if(r.getCell(CSGCodeHeaderIndex)== null)
                {
                }
                else
                CSGCode = cellToString(r.getCell(CSGCodeHeaderIndex));
                
                if(r.getCell(DoNotConvertHeaderIndex)== null)
                {
                }
                else
               	DoNotConvert = cellToString(r.getCell(DoNotConvertHeaderIndex));
                
               
                if(r.getCell(ProvisionedHeaderIndex)== null)
                {
                }
                else
                Provisioned = cellToString(r.getCell(ProvisionedHeaderIndex));
                
                
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
                newRow.createCell(cellIndex++).setCellValue(Category);
                newRow.createCell(cellIndex++).setCellValue(CSGCode);
                newRow.createCell(cellIndex++).setCellValue(DoNotConvert);
                newRow.createCell(cellIndex++).setCellValue(Provisioned);
                i++;
                }
           // }
  			  
            }
            
            FileOutputStream fos = new FileOutputStream(new File("D:\\backup\\GC\\GC_RES3.xlsx"));
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

        


   