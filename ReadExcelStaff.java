package task13;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcelStaff {

	public static void main(String[] args) {

		File ReadExcel= new File("D:\\Guvi Task\\Task 13\\Staff.xlsx");
		FileInputStream fis = null;
		XSSFWorkbook wb=null;
		XSSFSheet sheet=null;
		try {
			fis = new FileInputStream(ReadExcel);
		} catch (FileNotFoundException e) {
			e.printStackTrace();		
		}
		
		try {
			 wb=new XSSFWorkbook(fis);
		} catch (IOException e) {
			e.printStackTrace();
		}
		
		 sheet=wb.getSheetAt(0);
		 
		 int totalRows=sheet.getLastRowNum()+1; 
		 int totalColumns=sheet.getRow(0).getLastCellNum();
		 
		 System.out.println("Total No of Rows: "+totalRows);
	    System.out.println("Total No of Columns: "+totalColumns);
		 System.out.println();
		
		 for(int currentRow=1; currentRow<totalRows; currentRow++) 
			{
			 System.out.println(sheet.getRow(currentRow).getCell(0).getStringCellValue());
			 System.out.println(sheet.getRow(currentRow).getCell(1).getStringCellValue());
			 System.out.println(sheet.getRow(currentRow).getCell(2).getStringCellValue());
			 System.out.println("\t");	
			} 
		 
	}

}
