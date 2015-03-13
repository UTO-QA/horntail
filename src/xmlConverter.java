import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class xmlConverter {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		// Read in the Excel file
		InputStream ExcelFileToRead = new FileInputStream("Returning_Grad.xlsx");
		// Make a .xlsx workbook from .xlsx file
		XSSFWorkbook workBook = new XSSFWorkbook(ExcelFileToRead);

		// Get the sheet from the .xlsx workbook
		XSSFSheet sheet = workBook.getSheetAt(0);

		// Declare variables for manipulation of .xlsx workbook
		XSSFRow row;
		XSSFCell cell;

		// Define Iterators for the different sheets
		Iterator<Row> rowIterator = sheet.rowIterator();

		// Skip the column headings and already executed rows
		row = (XSSFRow) rowIterator.next();
		row = (XSSFRow) rowIterator.next();

		// Initialize the cell variable
		Iterator<Cell> cells = row.cellIterator();
		cell = (XSSFCell) cells.next();

		BufferedWriter writer = null;
		File xmlFile = new File("xmlFile.xml");
		writer = new BufferedWriter(new FileWriter(xmlFile, true));
		writer.write("<?xml version=\"1.0\"?>\n");
		writer.write("<testdata>\n");
		while (rowIterator.hasNext()) {
			writer.write("<vars FirstName=\"" + cell.getStringCellValue()
					+ "\" ");
			if (cells.hasNext()) {
				cell = (XSSFCell) cells.next();
			}
			writer.write("LastName=\"" + cell.getStringCellValue() + "\" ");
			if (cells.hasNext()) {
				cell = (XSSFCell) cells.next();
			}
			writer.write("Email=\"" + cell.getStringCellValue() + "\" ");
			if (cells.hasNext()) {
				cell = (XSSFCell) cells.next();
				if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
					String temp = cell.getStringCellValue().replaceAll("\\s+","");
					temp = temp.replaceAll("[(]","");
					temp = temp.replaceAll("[)]","-");
					writer.write("Phone=\"" + temp
							+ "\"/>\n");
				} else {
					writer.write("Phone=\"" + cell.getNumericCellValue()
							+ "\"/>\n");
				}
			} else {
				writer.write("Phone=\"" + "\"/>\n");
			}
			row = (XSSFRow) rowIterator.next();
			cells = row.cellIterator();
			if (cells.hasNext()) {
				cell = (XSSFCell) cells.next();
			}
		}
		writer.append("</testdata>\n");
		System.out.println(xmlFile.getCanonicalPath());
		writer.close();

	}
}
