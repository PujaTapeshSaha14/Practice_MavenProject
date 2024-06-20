// Author Puja Saha
// Date of creation: 22-06-2024
package ExcelReadWrite;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteInExcel {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub

		
		FileOutputStream file = new FileOutputStream("C:\\Users\\MT942UU\\eclipse-workspace\\MavenProject\\src\\test\\resources\\testdata\\myfile.xlsx");
		
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Data");
		XSSFRow row = sheet.createRow(0);
		row.createCell(0).setCellValue("Puja");
		row.createCell(1).setCellValue("BTech");
		row.createCell(2).setCellValue("ITER");
		
		workbook.write(file);
		
		workbook.close();
		file.close();
	}

}
