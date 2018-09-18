package UtilFiles;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadAndWriteExcelData {
	

	XSSFWorkbook xb;
	XSSFSheet Sheet1;
	String FilePath = "C:\\Users\\mgopalan\\Desktop\\VSTSREDMINE.xlsx";
	public ReadAndWriteExcelData(String FilePath){
		
		try {
			File src=new File(FilePath);
			FileInputStream FIS=new FileInputStream(src);
            xb = new XSSFWorkbook(FIS);
            
			
		
		} catch (Exception e) {
			// TODO Auto-generated catch block
			System.out.println(e.getMessage());
		}	
		
		    
	}
	
public String getData(int SheetNum, int row,int column)
	{   
		String data = " ";
		Sheet1=xb.getSheetAt(SheetNum);
		Cell cell = Sheet1.getRow(row).getCell(column);
		
	    
	    //String data=Sheet1.getRow(row).getCell(column).getStringCellValue();
		
		if (cell == null) {
			   data = "no data";
		} else if (cell.getCellTypeEnum()== CellType.STRING) {
			
			data = cell.getStringCellValue();
			//Date date = cell.getDateCellValue();
			
		}
		else if (cell.getCellTypeEnum() == CellType.NUMERIC){
			
			if(DateUtil.isCellDateFormatted(cell)) {
				
				Date date = cell.getDateCellValue();
				data = date.toString();
			}
			else {
				int a  = (int) cell.getNumericCellValue();
				data = Integer.toString(a);
			}
		 }
		
		return data;
		
	}
	

	
public void writeData(int SheetNum, int rowNum, int column, String vstsID) throws IOException {
		
		
		File src=new File(FilePath);
		FileOutputStream outputStream = new FileOutputStream(src);
	
		Sheet1=xb.getSheetAt(SheetNum);
		Row row =Sheet1.getRow(rowNum);
		
		
		Cell cell = row.createCell(column);
		cell.setCellValue(vstsID);
	
		System.out.println(cell.getStringCellValue());
		xb.write(outputStream);
		
		outputStream.close();
		
	}


}
