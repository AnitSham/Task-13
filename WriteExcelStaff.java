package task13;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteExcelStaff {

	public static void main(String[] args) throws IOException {

		XSSFWorkbook wb = null;
		XSSFSheet sheet = null;
		FileOutputStream fos = null;

		wb = new XSSFWorkbook();
		sheet = wb.createSheet("sheet1");

		Object[][] data = { { "Name", "Age", "Email" }, { "John Doe", "30", "john@test.com" },
				{ "Jane Doe", "28", "john@test.com" }, { "Bob Smith", "35", "jacky@example.com" },
				{ "Swapnil", "37", "swapnil@example.com" } };

		int rowcount = 0;
		for (Object[] row1 : data) {
			XSSFRow row = sheet.createRow(rowcount++);
			int columncount = 0;
			for (Object col : row1) {
				XSSFCell cell = row.createCell(columncount++);

				if (col instanceof String) {
					cell.setCellValue((String) col);
				} else if (col instanceof Integer) {
					cell.setCellValue((Integer) col);
				}
			}
		}
		fos = new FileOutputStream("D:\\Guvi Task\\Task 13\\StaffWrite.xlsx");
		wb.write(fos);
		fos.close();
	}
}
