package file;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class FileOperations {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		FileOperations obj = new FileOperations();
		try {
			obj.writeExcel();
			obj.readExcel();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
	}
	
	public void writeExcel() throws IOException {
		// TODO Auto-generated method stub
		File file=new File(System.getProperty("user.dir")+"\\src\\main\\resources\\book.xlsx");
		FileOutputStream output = new FileOutputStream(file);
		XSSFWorkbook book =new XSSFWorkbook();
		XSSFSheet sheet =book.createSheet("Sheet1");
		sheet.createRow(0).createCell(0).setCellValue("Name");
		sheet.getRow(0).createCell(1).setCellValue("Age");
		sheet.getRow(0).createCell(2).setCellValue("Email");
		
		sheet.createRow(1).createCell(0).setCellValue("John Doe");
		sheet.getRow(1).createCell(1).setCellValue("30");
		sheet.getRow(1).createCell(2).setCellValue("john@test.com");

		sheet.createRow(2).createCell(0).setCellValue("Jane Doe");
		sheet.getRow(2).createCell(1).setCellValue("28");
		sheet.getRow(2).createCell(2).setCellValue("jane@test.com");
		
		sheet.createRow(3).createCell(0).setCellValue("Bob Smith");
		sheet.getRow(3).createCell(1).setCellValue("35");
		sheet.getRow(3).createCell(2).setCellValue("jackey@example.com");
		
		sheet.createRow(4).createCell(0).setCellValue("Swapnil");
		sheet.getRow(4).createCell(1).setCellValue("37");
		sheet.getRow(4).createCell(2).setCellValue("swapnil@example.com");
		
		book.write(output);
		
		book.close();
		output.close();
	}
	
	public void readExcel() throws IOException {
		File file=new File(System.getProperty("user.dir")+"\\src\\main\\resources\\book.xlsx");
		FileInputStream input = new FileInputStream(file);
		XSSFWorkbook book =new XSSFWorkbook(input);
		XSSFSheet sheet =book.getSheet("Sheet1");
		int lastRow=sheet.getLastRowNum();
		for(int i=0; i<=lastRow; i++) {
			for(int j=0; j<=2; j++) {
				Cell cell= sheet.getRow(i).getCell(j);
				DataFormatter format = new DataFormatter();
				String s = format.formatCellValue(cell);
				System.out.print(s+" ");
			}
			System.out.println();
		}
		book.close();
		input.close();
	}

}

Output:
Name Age Email 
John Doe 30 john@test.com 
Jane Doe 28 jane@test.com 
Bob Smith 35 jackey@example.com 
Swapnil 37 swapnil@example.com 

