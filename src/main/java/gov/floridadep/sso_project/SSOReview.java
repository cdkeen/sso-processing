package gov.floridadep.sso_project;

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
 * Program reads an SSO check sheet and generates an email update for case managers.
 * Only works for .xlsx files with a specific database schema.
 * Code will need to be adapted to any database schema changes.
 * Make sure SSO data is consistent, otherwise code will break!
 * 
 */

public class SSOReview {

	public static void main(String[] args) throws IOException {
		readExcel();
	}
	//Method reads .xlsx file
	static void readExcel() throws IOException {
		//Get path to SSO database .xlsx file to be read
		String excelFilePath = "D:\\FDEP\\SSO_project\\SSO_database.xlsx";
		//obtaining input bytes from file
		FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
		//Create Workbook instance holding reference to file
		Workbook workbook = new XSSFWorkbook(inputStream);
		//Get the Sheet object at the given index
		Sheet firstSheet = workbook.getSheetAt(0);
		//Returns an iterator of the sheets in the workbook in sheet order
/*		Iterator<Row> iterator = firstSheet.iterator();
				
		//Iterate through each row one by one
		while (iterator.hasNext()) {
		Row nextRow = iterator.next();
		//For each row, iterate through all the columns
		Iterator<Cell> cellIterator = nextRow.cellIterator();
					
			while (cellIterator.hasNext()) {
				Cell cell = cellIterator.next();
				//Check the cell type and format accordingly		
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
					case BLANK:
						System.out.print("NULL");
				default:
						System.out.print("_");
					break;
				}
				System.out.print(" ");
			}
			System.out.println();
		}
		// Get the row object at the given index
*/		Row row = firstSheet.getRow(1);
		// Get cell object at the given index
		Cell cell = row.getCell(33);
			
		

		//for rows that contain "Jenny" in the "case manager" column
			//get needed data
		
		workbook.close();
		inputStream.close();		
	}
}
