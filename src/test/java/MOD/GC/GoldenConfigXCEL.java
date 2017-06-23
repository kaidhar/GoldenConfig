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

public class GoldenConfigXCEL {
	@Test
     public static void main() throws IOException, InvalidFormatException{

         

		
		InputStream is = new FileInputStream(new File("D:\\backup\\GC\\WVHeartland.xlsx"));
			
		Workbook workbook = StreamingReader.builder()
		        .rowCacheSize(100)    // number of rows to keep in memory (defaults to 10)
		        .bufferSize(4096)     // buffer size to use when reading InputStream to file (defaults to 1024)
		        .open(is);            // InputStream or File for XLSX file (required)
		Sheet sheet = workbook.getSheetAt(4);
		for (Row r : sheet) {
			  for (Cell c : r) {
				  
				  
				  
				  
			    System.out.println(c.getStringCellValue());
			  }
			}     
        
	
    
          	long rowNum = 10000;
            //int rowNum = sheet.getLastRowNum() + 1;
            //int colNum = sheet.getRow(0).getLastCellNum();
            int colNum = 20;
            int CORPHeaderIndex = -1, DDPHeaderIndex = -1;
            int CSGHeaderIndex = -1, RATEHeaderIndex = -1;
            
            Row rowHeader = sheet.getRow(0);
            for (int j = 0; j < colNum; j++) {
            	Cell cell = rowHeader.getCell(j);
                String cellValue = cellToString(cell);
                if ("CORP".equalsIgnoreCase(cellValue)) {
                	CORPHeaderIndex = j;
                } else if ("DDP".equalsIgnoreCase(cellValue)) {
                	DDPHeaderIndex = j;
                }else if ("RATE".equalsIgnoreCase(cellValue)) {
                	CSGHeaderIndex = j;
                }else if ("CSG".equalsIgnoreCase(cellValue)) {
                	RATEHeaderIndex = j;
                    
                
            }
            }
            if (RATEHeaderIndex == -1 || CSGHeaderIndex == -1) {
                System.out.println("LOLA");
            }
            
         // createnew workbook
            XSSFWorkbook workbook1 = new XSSFWorkbook();
            // Create a blank sheet
            XSSFSheet sheet1 = workbook1.createSheet("DATA");
            XSSFRow newRow = sheet1.createRow(0); 
            int HeadIndex = 0;
            newRow.createCell(HeadIndex++).setCellValue("CORP");
            newRow.createCell(HeadIndex++).setCellValue("DDP");
            newRow.createCell(HeadIndex++).setCellValue("RATE");
            newRow.createCell(HeadIndex++).setCellValue("CSG");
            
            
            for (int i = 1; i < rowNum; i++) {
                Row row = sheet.getRow(i);
                
                String CORP = cellToString(row.getCell(CORPHeaderIndex));
                String DDP = cellToString(row.getCell(DDPHeaderIndex));
                String RATE = cellToString(row.getCell(RATEHeaderIndex));
                String CSG = cellToString(row.getCell(CSGHeaderIndex));
                int cellIndex = 0;
                if(CORP.contains("99938"))
                {
                //Create a newRow object for the output excel. 
                //We begin for i = 1, because of the headers from the input excel, so we go minus 1 in the new (no headers).
                //If for the output we need headers, add them outside this for loop, and go with i, not i-1
                newRow = sheet1.createRow(i);  
                newRow.createCell(cellIndex++).setCellValue(CORP);
                newRow.createCell(cellIndex++).setCellValue(DDP);
                newRow.createCell(cellIndex++).setCellValue(RATE);
                newRow.createCell(cellIndex++).setCellValue(CSG);
                }
            }
            
            FileOutputStream fos = new FileOutputStream(new File("test1.xlsx"));
            workbook.write(fos);
            fos.close();
            
            }
	
	
            public static String cellToString(Cell cell) 
            {
                int type;
                Object result = null;
                type = cell.getCellType();

                switch (type) {

                case XSSFCell.CELL_TYPE_NUMERIC:
                    result = BigDecimal.valueOf(cell.getNumericCellValue())
                            .toPlainString();

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

        


   