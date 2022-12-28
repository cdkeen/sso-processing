package gov.floridadep.sso_project;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/*
 * Program reads an SSO excel check sheet and prints a spill update for each case manager to the console. 
 */

public class SSOUpdate {

	public static void main(String[] args) throws IOException {
		readExcel();
	}
	// Method reads .xlsx file and prints specified cell values
	static void readExcel() throws IOException {
		/* 
		 * Work laptop file path: "C:\\Users\\Keen_CD\\OneDrive - Florida Department of Environmental Protection\\Desktop\\SSO\\Data_entry_project\\SSO_database.xlsx"
		 * Home PC File path: "D:\\FDEP\\SSO_project\\SSO_database.xlsx"
		 * Macbook file path: "/Users/cdkeen/Documents/FDEP/sso-project/SSO_database.xlsx"
		 */
		String excelFilePath = "/Users/cdkeen/Documents/FDEP/sso-project/SSO_database.xlsx";
		FileInputStream inputStream = new FileInputStream(new File(excelFilePath)); // Obtaining input bytes from file
		Workbook workbook = new XSSFWorkbook(inputStream); // Create Workbook instance holding reference to file
		Sheet firstSheet = workbook.getSheetAt(0); // Get the Sheet object at the given index (returns first sheet in the workbook)
	/*	
		Scanner scanner = new Scanner(System.in); // User inputs the starting row number
		System.out.println("Enter starting row number:");
		while (!scanner.hasNextInt()) {
			System.out.println("Invalid entry, please enter a row number");
			scanner.next();
		}
		int rowIndex = scanner.nextInt() - 1;
	*/	
	//	Iterator<Row> rowIterator = firstSheet.iterator(); // Get an iterator for the sheet
	//	Row row = rowIterator.next();
		
		DataFormatter formatter = new DataFormatter(); // Fetch the string value of the case manager cell
		//Loop thru the sheet for each case manager:
		System.out.println("Carolyn");
		for (Row row : firstSheet) {
	    	String caseManagerName = formatter.formatCellValue(row.getCell(33));
	    	Cell dateCell = row.getCell(5);
			Cell letterCell = row.getCell(28);
			Cell spillTypeCell = row.getCell(15);
			Cell spillLocationCell = row.getCell(9);
			Cell facilityNameCell = row.getCell(3);
			Cell statusCell = row.getCell(35);
			
			if (caseManagerName.contains("Carolyn")) {
				System.out.print(dateCell + " - ");
				System.out.print(letterCell + " - ");
				System.out.print(spillTypeCell + " - ");
				System.out.print(spillLocationCell + " - ");
				System.out.print(facilityNameCell + " - ");
				System.out.print(statusCell);
				System.out.println();
			}
		}
		System.out.println();
		System.out.println("Jenny");
		for (Row row : firstSheet) {
	    	String caseManagerName = formatter.formatCellValue(row.getCell(33));
	    	Cell dateCell = row.getCell(5);
			Cell letterCell = row.getCell(28);
			Cell spillTypeCell = row.getCell(15);
			Cell spillLocationCell = row.getCell(9);
			Cell facilityNameCell = row.getCell(3);
			Cell statusCell = row.getCell(35);
			
			if (caseManagerName.contains("Jenny")) {
				System.out.print(dateCell + " - ");
				System.out.print(letterCell + " - ");
				System.out.print(spillTypeCell + " - ");
				System.out.print(spillLocationCell + " - ");
				System.out.print(facilityNameCell + " - ");
				System.out.print(statusCell);
				System.out.println();
			}
		}
		System.out.println();
		System.out.println("Amanda");
		for (Row row : firstSheet) {
	    	String caseManagerName = formatter.formatCellValue(row.getCell(33));
	    	Cell dateCell = row.getCell(5);
			Cell letterCell = row.getCell(28);
			Cell spillTypeCell = row.getCell(15);
			Cell spillLocationCell = row.getCell(9);
			Cell facilityNameCell = row.getCell(3);
			Cell statusCell = row.getCell(35);
			
			if (caseManagerName.contains("Amanda")) {
				System.out.print(dateCell + " - ");
				System.out.print(letterCell + " - ");
				System.out.print(spillTypeCell + " - ");
				System.out.print(spillLocationCell + " - ");
				System.out.print(facilityNameCell + " - ");
				System.out.print(statusCell);
				System.out.println();
			}
		}
		System.out.println();
		System.out.println("Gina");
		for (Row row : firstSheet) {
	    	String caseManagerName = formatter.formatCellValue(row.getCell(33));
	    	Cell dateCell = row.getCell(5);
			Cell letterCell = row.getCell(28);
			Cell spillTypeCell = row.getCell(15);
			Cell spillLocationCell = row.getCell(9);
			Cell facilityNameCell = row.getCell(3);
			Cell statusCell = row.getCell(35);
			
			if (caseManagerName.contains("Gina")) {
				System.out.print(dateCell + " - ");
				System.out.print(letterCell + " - ");
				System.out.print(spillTypeCell + " - ");
				System.out.print(spillLocationCell + " - ");
				System.out.print(facilityNameCell + " - ");
				System.out.print(statusCell);
				System.out.println();
			}
		}
		System.out.println();
		System.out.println("Manny");
		for (Row row : firstSheet) {
	    	String caseManagerName = formatter.formatCellValue(row.getCell(33));
	    	Cell dateCell = row.getCell(5);
			Cell letterCell = row.getCell(28);
			Cell spillTypeCell = row.getCell(15);
			Cell spillLocationCell = row.getCell(9);
			Cell facilityNameCell = row.getCell(3);
			Cell statusCell = row.getCell(35);
			
			if (caseManagerName.contains("Manny")) {
				System.out.print(dateCell + " - ");
				System.out.print(letterCell + " - ");
				System.out.print(spillTypeCell + " - ");
				System.out.print(spillLocationCell + " - ");
				System.out.print(facilityNameCell + " - ");
				System.out.print(statusCell);
				System.out.println();
			}
		}
		System.out.println();
		System.out.println("Hannah");
		for (Row row : firstSheet) {
	    	String caseManagerName = formatter.formatCellValue(row.getCell(33));
	    	Cell dateCell = row.getCell(5);
			Cell letterCell = row.getCell(28);
			Cell spillTypeCell = row.getCell(15);
			Cell spillLocationCell = row.getCell(9);
			Cell facilityNameCell = row.getCell(3);
			Cell statusCell = row.getCell(35);
			
			if (caseManagerName.contains("Hannah")) {
				System.out.print(dateCell + " - ");
				System.out.print(letterCell + " - ");
				System.out.print(spillTypeCell + " - ");
				System.out.print(spillLocationCell + " - ");
				System.out.print(facilityNameCell + " - ");
				System.out.print(statusCell);
				System.out.println();
			}
		}
		System.out.println();
		System.out.println("Cory");
		for (Row row : firstSheet) {
	    	String caseManagerName = formatter.formatCellValue(row.getCell(33));
	    	Cell dateCell = row.getCell(5);
			Cell letterCell = row.getCell(28);
			Cell spillTypeCell = row.getCell(15);
			Cell spillLocationCell = row.getCell(9);
			Cell facilityNameCell = row.getCell(3);
			Cell statusCell = row.getCell(35);
			
			if (caseManagerName.contains("Cory")) {
				System.out.print(dateCell + " - ");
				System.out.print(letterCell + " - ");
				System.out.print(spillTypeCell + " - ");
				System.out.print(spillLocationCell + " - ");
				System.out.print(facilityNameCell + " - ");
				System.out.print(statusCell);
				System.out.println();
			}
		}
		System.out.println();
		System.out.println("Trey");
		for (Row row : firstSheet) {
	    	String caseManagerName = formatter.formatCellValue(row.getCell(33));
	    	Cell dateCell = row.getCell(5);
			Cell letterCell = row.getCell(28);
			Cell spillTypeCell = row.getCell(15);
			Cell spillLocationCell = row.getCell(9);
			Cell facilityNameCell = row.getCell(3);
			Cell statusCell = row.getCell(35);
			
			if (caseManagerName.contains("Trey")) {
				System.out.print(dateCell + " - ");
				System.out.print(letterCell + " - ");
				System.out.print(spillTypeCell + " - ");
				System.out.print(spillLocationCell + " - ");
				System.out.print(facilityNameCell + " - ");
				System.out.print(statusCell);
				System.out.println();
			}
		}
		System.out.println();
		System.out.println("Sean");
		for (Row row : firstSheet) {
	    	String caseManagerName = formatter.formatCellValue(row.getCell(33));
	    	Cell dateCell = row.getCell(5);
			Cell letterCell = row.getCell(28);
			Cell spillTypeCell = row.getCell(15);
			Cell spillLocationCell = row.getCell(9);
			Cell facilityNameCell = row.getCell(3);
			Cell statusCell = row.getCell(35);
			
			if (caseManagerName.contains("Sean")) {
				System.out.print(dateCell + " - ");
				System.out.print(letterCell + " - ");
				System.out.print(spillTypeCell + " - ");
				System.out.print(spillLocationCell + " - ");
				System.out.print(facilityNameCell + " - ");
				System.out.print(statusCell);
				System.out.println();
			}
		}
			
		/*				
			
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
			
		*/	
		workbook.close();
		inputStream.close();
	}
}
			
