package gov.fdep.cd.sso_review;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/*
 * Program reads an SSO database spreadsheet and prints to the console (skips empty cells).
 */

public class SSOReview {

	public static void main(String[] args) throws IOException {
		readExcel();
	}
	//Method reads .xlsx file
	static void readExcel() throws IOException {
		//Get SSO database excel file
		String excelFilePath = "C:\\Users\\Keen_CD\\OneDrive - Florida Department of Environmental Protection\\Desktop\\SSO\\Data_entry_project\\SSO_database.xlsx";
		//obtaining input bytes from a file
		FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
		//creating workbook instance that refers to .xlsx file
		Workbook workbook = new XSSFWorkbook(inputStream);
		//Get the Sheet object at the given index
		Sheet firstSheet = workbook.getSheetAt(0);
		//Returns an iterator of the sheets in the workbook in sheet order
		Iterator<Row> iterator = firstSheet.iterator();
				
		//If this scanner has another token in its input, run the while loop
		while (iterator.hasNext()) {
		Row nextRow = iterator.next();
		Iterator<Cell> cellIterator = nextRow.cellIterator();
					
			while (cellIterator.hasNext()) {
				Cell cell = cellIterator.next();
						
				switch (cell.getCellType()) {
					case STRING:
						System.out.print(cell.getStringCellValue());
						break;
					case BOOLEAN:
						System.out.print(cell.getBooleanCellValue());
						break;
					case NUMERIC:
						System.out.print(cell.getNumericCellValue());
						break;
				default:
					break;
				}
				System.out.print(" | ");
			}
			System.out.println();
		}
		
		workbook.close();
		inputStream.close();		
	}
}
