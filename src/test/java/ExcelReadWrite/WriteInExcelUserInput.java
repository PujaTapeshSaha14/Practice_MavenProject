package ExcelReadWrite;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Scanner;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteInExcelUserInput {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub

		
		FileOutputStream file = new FileOutputStream("C:\\Users\\MT942UU\\eclipse-workspace\\MavenProject\\src\\test\\resources\\testdata\\myfile1.xlsx");
		
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Data1");
		
		Scanner sc = new Scanner(System.in);
		System.out.println("Enter number of row: ");
		int noOfRows=sc.nextInt();
		System.out.println("Enter number of cells: ");
		int noOfCells=sc.nextInt();
		for(int r=0; r<=noOfRows; r++)
		{
			XSSFRow row=sheet.createRow(r);
			for(int c=0; c<noOfCells; c++)
			{
				XSSFCell cell=row.createCell(c);
				cell.setCellValue(sc.next());
			}
		}
	
		workbook.write(file);
		System.out.println("File is created");
		workbook.close();
		file.close();
	}

}
