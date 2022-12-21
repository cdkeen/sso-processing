package gov.floridadep.sso_project;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.Scanner;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/*
 * Program reads an SSO check sheet and generates an email update for case managers. 
 */

public class SSOUpdate {

	public static void main(String[] args) throws IOException {
		
		readExcel();
	}
	// Method reads .xlsx file
	static void readExcel() throws IOException {
		// Get path to SSO database .xlsx file
		// Work laptop file path: "C:\\Users\\Keen_CD\\OneDrive - Florida Department of Environmental Protection\\Desktop\\SSO\\Data_entry_project\\SSO_database.xlsx"
		// Home PC File path: "D:\\FDEP\\SSO_project\\SSO_database.xlsx"
		String excelFilePath = "C:\\Users\\Keen_CD\\OneDrive - Florida Department of Environmental Protection\\Desktop\\SSO\\Data_entry_project\\SSO_database.xlsx";
		// Obtaining input bytes from file
		FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
		// Create Workbook instance holding reference to file
		Workbook workbook = new XSSFWorkbook(inputStream);
		// Get the Sheet object at the given index (returns first sheet in the workbook)
		Sheet firstSheet = workbook.getSheetAt(0);
		
		// User inputs the starting row number
		// The update will include this row and all the following rows until the end of the sheet
		Scanner scanner = new Scanner(System.in);
		System.out.println("Enter starting row number:");
		while (!scanner.hasNextInt()) {
			System.out.println("Invalid entry, please enter a row number");
			scanner.next();
		}
		Iterator<Row> rowIterator = firstSheet.iterator();
		//Parent for loop
	//	for (int rowIndex = scanner.nextInt() - 1; rowIndex <= firstSheet.getLastRowNum(); rowIndex++) {
	//		Row nextRow = rowIterator.next();
	//	}
		
		
		
/*		for (Sheet sheet : wb ) {
		    for (Row row : sheet) {
		        for (Cell cell : row) {
		            // Do something here
		        }
		    }
		}
*/	
		
			
		/*	// Returns an iterator for the sheet
		Iterator<Row> rowIterator = firstSheet.iterator();			
			
		// For each row (starting at the row specified by the user), do this:
		for (int rowIndex = scanner.nextInt() - 1; rowIndex <= firstSheet.getLastRowNum(); rowIndex++) {
			// Return the next row in the iteration
			Row nextRow = rowIterator.next();
			// Return a cell iterator for the row
			Iterator<Cell> cellIterator = nextRow.cellIterator();
			// Get the row object at the given index
			Row row = firstSheet.getRow(rowIndex);
			// Get case manager cell for the row
			Cell caseManagerCell = row.getCell(33);
			// Fetch the string value of the case manager cell
			DataFormatter formatter = new DataFormatter();
			String caseManager = formatter.formatCellValue(caseManagerCell);
			// Check for who the case manager is for the row and act accordingly:
			switch (caseManager) {
				case "Carolyn - Brevard, City of Winter Park":
					
					Cell dateCell = row.getCell(5);
					Cell letterCell = row.getCell(28);
					Cell spillTypeCell = row.getCell(15);
					Cell spillLocationCell = row.getCell(9);
					Cell facilityNameCell = row.getCell(3);
					Cell statusCell = row.getCell(35);	
					
					System.out.print(dateCell + " - ");
					System.out.print(letterCell + " - ");
					System.out.print(spillTypeCell + " - ");
					System.out.print(spillLocationCell + " - ");
					System.out.print(facilityNameCell + " - ");
					System.out.print(statusCell);
					System.out.println();
					break;			
			}
			
		*/	
		scanner.close();	
		}
		
		
	//	workbook.close();
		//inputStream.close();
		}
			
