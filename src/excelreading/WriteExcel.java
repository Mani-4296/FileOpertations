package excelreading;

	import java.io.File;
	import java.io.FileOutputStream;
	import java.util.Scanner;

	import org.apache.poi.xssf.usermodel.XSSFCell;
	import org.apache.poi.xssf.usermodel.XSSFRow;
	import org.apache.poi.xssf.usermodel.XSSFSheet;
	import org.apache.poi.xssf.usermodel.XSSFWorkbook;

	public class WriteExcel {

	    public static void main(String[] args) {
	        WriteExcel obj = new WriteExcel();
	        obj.writeExcel();
	        System.out.println("Data has been written to Excel successfully!");
	    }

	    public void writeExcel() {
	        try {
	            Scanner scanner = new Scanner(System.in);
	            System.out.println("Enter the file path to save the Excel file:");
	            String filePath = scanner.nextLine();

	            XSSFWorkbook wb = new XSSFWorkbook(); // Create a new Excel workbook
	            XSSFSheet sheet = wb.createSheet("Sheet1"); // Create a new sheet in the workbook

	            System.out.println("Enter the number of rows:");
	            int numRows = scanner.nextInt();
	            scanner.nextLine(); // Consume newline

	            for (int i = 0; i < numRows; i++) {
	                System.out.println("Enter data for row " + (i + 1) + " (comma separated):");
	                String rowData = scanner.nextLine();
	                String[] cellData = rowData.split(",");

	                XSSFRow row = sheet.createRow(i); // Create a new row in the sheet
	                for (int j = 0; j < cellData.length; j++) {
	                    XSSFCell cell = row.createCell(j); // Create a new cell in the row
	                    cell.setCellValue(cellData[j]); // Set the cell value
	                }
	            }

	            scanner.close();

	            FileOutputStream fos = new FileOutputStream(new File(filePath)); // Create a file output stream for the Excel file
	            wb.write(fos); // Write the workbook data to the file
	            wb.close(); // Close the workbook
	            fos.close(); // Close the file output stream
	        } catch (Exception e) {
	            System.out.println("An error occurred while writing data to Excel:");
	            e.printStackTrace();
	        }
	    }
	}

