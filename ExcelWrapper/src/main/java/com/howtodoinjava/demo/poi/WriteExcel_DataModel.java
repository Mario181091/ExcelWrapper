package com.howtodoinjava.demo.poi;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Iterator;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;

public class WriteExcel_DataModel {

	public static int numrow = 0;
	public static int numTable = 0;
	public static int numberOfCellSummaryRow = 12;
	public static int numberOfCellDefSheets = 13;

	public static void main(String[] args) {
		UtilityFunctions utilityFunctions = new UtilityFunctions(numberOfCellSummaryRow, numberOfCellDefSheets);
		CellStyles cellStyles = new CellStyles();

		try {
			startComputing(utilityFunctions, cellStyles);
		} catch (EncryptedDocumentException | InvalidFormatException | IOException e1) {
			e1.printStackTrace();
		}

		System.out.println();
		System.out.println("************************************************************************************");
		System.out.println("       .-\"-.            .-\"-.            .-\"-.           .-\"-.\r\n"
				+ "     _/_-.-_\\_        _/.-.-.\\_        _/.-.-.\\_       _/.-.-.\\_\r\n"
				+ "    / __} {__ \\      /|( o o )|\\      ( ( o o ) )     ( ( o o ) )\r\n"
				+ "   / //  \"  \\\\ \\    | //  \"  \\\\ |      |/  \"  \\|       |/  \"  \\|\r\n"
				+ "  / / \\'---'/ \\ \\  / / \\'---'/ \\ \\      \\'/^\\'/         \\ .-. /\r\n"
				+ "  \\ \\_/`\"\"\"`\\_/ /  \\ \\_/`\"\"\"`\\_/ /      /`\\ /`\\         /`\"\"\"`\\\r\n"
				+ "   \\           /    \\           /      /  /|\\  \\       /       \\\r\n" + "\r\n"
				+ " -={ see no GP } = { hear no GP } = { speak no GP } = { have no fun }=-");
	}

	@SuppressWarnings("deprecation")
	public static void startComputing(UtilityFunctions utilityFunctions, CellStyles cellStyles)
			throws EncryptedDocumentException, InvalidFormatException, IOException {

		// utility fields
		String newSheet = "";
		String sheetName = "";
		int iter = 0;
		ArrayList<String> curretRowArray = new ArrayList<String>();
		ArrayList<String> listOfSheet = new ArrayList<String>();
		XSSFSheet sheet1 = null;

		// manage input file
		System.out.println("* Read input file 					****");
		FileInputStream fileInput = new FileInputStream(
				"C:\\Users\\m.carricato\\Desktop\\Project_OSS\\provaJava\\test_excel.xlsx");
		Workbook wb = WorkbookFactory.create(fileInput);
		Sheet sheet = wb.getSheet("Sheet1");
		Iterator<Row> rowIterator = sheet.rowIterator();

		// manage output file
		System.out.println("** Open output file  					*******");
		XSSFWorkbook wb_out = new XSSFWorkbook();
		OutputStream fileOut = new FileOutputStream(
				"C:\\Users\\m.carricato\\Desktop\\Project_OSS\\provaJava\\data_model_output.xlsx");

		// generate summary sheet
		System.out.println("*** Generate Summary Sheet				**************");
		utilityFunctions.generateSummarySheet(wb_out, cellStyles);

		// Read input file
		System.out.println("**** Start Core Algorithm				*******************");
		while (rowIterator.hasNext()) {
			Row row = rowIterator.next();
			Iterator<Cell> cellIterator = row.cellIterator();
			String currCellValue = "";
			iter = -1;
			while (cellIterator.hasNext()) {
				Cell cell = cellIterator.next();
				switch (cell.getCellType()) {
				case Cell.CELL_TYPE_STRING:
					currCellValue = cell.getStringCellValue();
					iter++;
					break;
				case Cell.CELL_TYPE_FORMULA:
					currCellValue = cell.getCellFormula();
					iter++;
					break;
				case Cell.CELL_TYPE_NUMERIC:
					if (DateUtil.isCellDateFormatted(cell)) {
						currCellValue = cell.getDateCellValue().toString();
					} else {
						currCellValue = Double.toString(cell.getNumericCellValue());
					}
					iter++;
					break;
				case Cell.CELL_TYPE_BLANK:
					currCellValue = "";
					iter++;
					break;
				case Cell.CELL_TYPE_BOOLEAN:
					currCellValue = Boolean.toString(cell.getBooleanCellValue());
					iter++;
					break;
				default:
					currCellValue = "";
					iter++;
					break;
				}
				if (iter == 0) {
					sheetName = currCellValue;
					if (!newSheet.equals(sheetName) && (!sheetName.equals(""))) {
						sheet1 = wb_out.createSheet("ent_" + sheetName.toLowerCase());
						newSheet = sheetName;
						listOfSheet.add(sheetName);
						numrow = 0;
					}
				}
				curretRowArray.add(currCellValue);
			}
			coreAlghoritm(newSheet, curretRowArray, wb_out, fileOut, sheet1, utilityFunctions, cellStyles);
			curretRowArray.clear();
		}
		System.out.println("*** End Core Algorithm 					***********************");
		System.out.println("** Fill Summary sheet  					**************************");
		utilityFunctions.fillSummarySheet(wb_out, listOfSheet, cellStyles);
		utilityFunctions.deleteDirtySheet(wb_out);

		System.out.println("* Starting autosizing  					****************************");
		utilityFunctions.autoSizeColumns(wb_out);
		wb_out.write(fileOut);
		fileInput.close();
	}

	private static void coreAlghoritm(String newSheet, ArrayList<String> tempStructure, XSSFWorkbook wb_out,OutputStream fileOut, XSSFSheet sheet1, 
			                           UtilityFunctions utilityFunctions, CellStyles cellStyles) throws IOException {

		numrow++;
		String sourceObj = "";
		String targetFiled = "";
		String targetFiledDataType = "";
		String nullable = "";
		String dataValue = "";
		String dataSourceType = "";

		XSSFRow rowhead = sheet1.createRow((short) 0);
		utilityFunctions.generateSheetsHeader(rowhead, wb_out, cellStyles);
		XSSFCellStyle def_style = cellStyles.createDefaultXSSFCellStyle(wb_out);
		ArrayList<XSSFCell> rowFile = new ArrayList<>();

		XSSFRow row = sheet1.createRow((short) numrow);
		for (int i = 0; i < 13; i++) {
			XSSFCell cell = row.createCell(i);
			cell.setCellStyle(def_style);
			rowFile.add(cell);
		}

		for (int i = 0; i < tempStructure.size(); i++) {
			String currValue = tempStructure.get(i).toString();
			dataValue = utilityFunctions.normalizeDataType(currValue).toLowerCase();
			targetFiled = utilityFunctions.computeTargetDataValue(tempStructure);
			if (!targetFiled.equals("")) {
				String[] parts = targetFiled.split("@");
				sourceObj = "ent_" + parts[0].toLowerCase();
				targetFiled = parts[1];
				targetFiledDataType = parts[2];
				nullable = parts[3];
			}
			row.getCell(i + 1).setCellValue(dataValue);
			row.getCell(7).setCellValue(sourceObj);
			row.getCell(8).setCellValue(targetFiled);
			row.getCell(9).setCellValue(targetFiledDataType);
			row.getCell(10).setCellValue(nullable);
			if (i == 2) {
				dataSourceType = utilityFunctions.manageDataSourceType(tempStructure, dataValue);
			}
		}
		row.getCell(3).setCellValue(dataSourceType);
		row.getCell(5).setCellValue("");
		row.getCell(6).setCellValue("");
	}
}