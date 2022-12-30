package gov.floridadep.sso_project;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.Scanner;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class AccessEntry {

	public static void main(String[] args) throws IOException{
		accessNotes();

	}
	// Method reads from excel sheet and prints an access entry description for the specified spill 
	static void accessNotes() throws IOException {
		String filePath = "D:\\FDEP\\SSO_project\\SSO_database.xlsx";
		FileInputStream inputStream = new FileInputStream(new File(filePath));
		Workbook workbook = new XSSFWorkbook(inputStream);
		Sheet firstSheet = workbook.getSheetAt(0);
		
		Scanner scanner = new Scanner(System.in); // User inputs the row number
		System.out.println("Enter row number for spill:");
		while (!scanner.hasNextInt()) {
			System.out.println("Invalid entry, please enter a row number");
			scanner.next();
		}
		int rowIndex = scanner.nextInt() - 1;
		Row row = firstSheet.getRow(rowIndex);
		Cell pnpCell = row.getCell(20); // PNP date
		Cell essaCell = row.getCell(1); // ESSA #
		Cell typeCell = row.getCell(15); // spill type
		Cell addressCell = row.getCell(9); // address
		Cell causeCell = row.getCell(22); // cause
		Cell volumeCell = row.getCell(13); // spill volume
		Cell recoveredCell = row.getCell(14); // volume recovered
		Cell cleanupCell = row.getCell(18); // cleanup actions
		
		System.out.println("PNP Received " + pnpCell);
		System.out.println("DEP Incident ID #: " + essaCell);
		
		DataFormatter formatter = new DataFormatter();
		String typeCellString = formatter.formatCellValue(row.getCell(33));
		if (typeCellString == "Fully Treated/Reclaimed") {
			System.out.print("Unauthorized discharge at ");
		}
		else {
			System.out.print("Spill at ");	
		}
		
		System.out.print(addressCell + " of " + typeCell + " wastewater due to " + causeCell);
		System.out.print(". Approximately " + volumeCell + " gallons released, " + recoveredCell + " recovered. " + cleanupCell);
		
		scanner.close();
	}
}
