package excelreading;

import java.io.File;
import java.io.FileInputStream;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcelData {

	public static void main(String[] args) {
		ReadExcelData obj = new ReadExcelData();
		String studata = obj.ReadExcel();
		System.out.println("Excel Data is: \n" + studata);
	}

	public String ReadExcel() {
		StringBuilder data = new StringBuilder(); // StringBuilder to store the read data

		try {
			String path = "D:\\Mani-Workings\\Guvi\\Java code\\FileOpertaions\\ReadExcelFile.xlsx";
			File file = new File(path);
			FileInputStream fis = new FileInputStream(file);
			XSSFWorkbook wb = new XSSFWorkbook(fis); // Load the Excel workbook
			XSSFSheet sheet = wb.getSheet("sheet1"); // Get the desired sheet from the workbook

			int rowCount = sheet.getPhysicalNumberOfRows(); // Get the number of rows in the sheet
			for (int i = 0; i < rowCount; i++) {
				XSSFRow row = sheet.getRow(i); // Get the current row
				int cellCount = row.getLastCellNum(); // Get the number of cells in the row
				for (int j = 0; j < cellCount; j++) {
					XSSFCell cell = row.getCell(j); // Get the current cell
					if (cell != null) {
						switch (cell.getCellType()) {
						case STRING:
							data.append(cell.getStringCellValue()).append("\t");
							break;
						case NUMERIC:
							data.append(cell.getNumericCellValue()).append("\t");
							break;
						default:
							data.append(" ").append("\t");
						}
					} else {
						data.append(" ").append("\t");
					}
				}
				data.append("\n"); // Append newline after each row
			}
		} catch (Exception e) {
			System.out.println("The excel file not found, enter correct filepath to locate");
			e.printStackTrace();
		}
		return data.toString(); // Return the read data as a string
	}
}
