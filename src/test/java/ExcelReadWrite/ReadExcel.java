package ExcelReadWrite;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
public class ReadExcel {





	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub

		FileInputStream file = new FileInputStream(System.getProperty("user.dir")+"\\\\src\\\\test\\\\resources\\testdata\\TestData.xlsx");
		//FileInputStream file = new FileInputStream("C:\\Users\\MT942UU\\eclipse-workspace\\MavenProject\\src\\main\\resources\\testdata\\\\TestData.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(file);

		XSSFSheet sheet=workbook.getSheet("Info");
		//workbook.getSheetAt(0);
		int totalRows=sheet.getLastRowNum();
		int totalCells = sheet.getRow(1).getLastCellNum();

		System.out.println("Number of Rows: "+ totalRows);
		System.out.println("Number of Cells: "+ totalCells);

		for(int r=0; r<=totalRows; r++)
		{
			XSSFRow currentRow=sheet.getRow(r);
			for(int c=0; c<totalCells; c++)
			{
				XSSFCell currentCell =currentRow.getCell(c);
				System.out.print(currentCell.toString()+ "\t");
			}
			System.out.println();
		}
		workbook.close();
		file.close();
	}

}
