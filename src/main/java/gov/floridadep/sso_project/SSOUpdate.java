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
		// Get path to SSO database .xlsx file to be read
		String excelFilePath = "D:\\FDEP\\SSO_project\\SSO_database.xlsx";
		// Obtaining input bytes from file
		FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
		// Create Workbook instance holding reference to file
		Workbook workbook = new XSSFWorkbook(inputStream);
		// Get the Sheet object at the given index
		Sheet firstSheet = workbook.getSheetAt(0);
		// Returns an iterator of the sheets in the workbook in sheet order
		Iterator<Row> rowIterator = firstSheet.iterator();
					
		// Ask user to input the row number that the program will start at
		Scanner scanner = new Scanner(System.in);
		System.out.println("Enter starting row number");
		while (!scanner.hasNextInt()) scanner.next();
		
		// For each row (starting at the row specified by the user), do this:
		for (int rowIndex = scanner.nextInt() - 1; rowIndex <= firstSheet.getLastRowNum(); rowIndex++) {
			// Return the next row in the iteration
			Row nextRow = rowIterator.next();
			// Return a cell iterator for the row
			Iterator<Cell> cellIterator = nextRow.cellIterator();
			// Get the row object at the given index
			Row row = firstSheet.getRow(rowIndex);
			// Get cells for the row
			Cell caseManagerCell = row.getCell(33);
			Cell dateCell = row.getCell(5);
			Cell letterCell = row.getCell(28);
			Cell spillTypeCell = row.getCell(15);
			Cell spillLocationCell = row.getCell(9);
			Cell facilityNameCell = row.getCell(3);
			Cell statusCell = row.getCell(35);
			// Fetch the string value of the case manager cell
			DataFormatter formatter = new DataFormatter();
			String caseManager = formatter.formatCellValue(caseManagerCell);
			// Check for who the case manager is for the row and act accordingly:
			switch (caseManager) {
				case "Carolyn":
					System.out.println(caseManagerCell);
					System.out.print(dateCell + " - ");
					System.out.print(letterCell + " - ");
					System.out.print(spillTypeCell + " - ");
					System.out.print(spillLocationCell + " - ");
					System.out.print(facilityNameCell + " - ");
					System.out.print(statusCell);
					System.out.println();
					break;
				
					
			}
			
		//	carolyn, manny, cory, hannah, trey, gina, amanda, shaun, jenny
			
		}
		
		scanner.close();
		workbook.close();
		inputStream.close();
		}
			
	}
