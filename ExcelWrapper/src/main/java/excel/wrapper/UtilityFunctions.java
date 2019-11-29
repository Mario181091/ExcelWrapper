package excel.wrapper;

import java.util.ArrayList;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class UtilityFunctions {

	int numberOfCellSummaryRow;
	int numberOfCellDefSheets;
	
	public UtilityFunctions(int numberOfCellSummaryRow,int numberOfCellDefSheets) {
		this.numberOfCellSummaryRow = numberOfCellSummaryRow;
		this.numberOfCellDefSheets  = numberOfCellDefSheets;
	};

@SuppressWarnings("deprecation")
public void generateSummarySheet(XSSFWorkbook wb_out, CellStyles cellStyles) {

	ArrayList<XSSFCell> headerFileSummary = new ArrayList<>();
	XSSFColor green = new XSSFColor(new java.awt.Color(0, 176, 80));
	XSSFColor blue = new XSSFColor(new java.awt.Color(84, 142, 212));
	XSSFCellStyle style_green = cellStyles.createHeaderXSSFCellStyle(wb_out, green);
	XSSFCellStyle style_blue = cellStyles.createHeaderXSSFCellStyle(wb_out, blue);
	style_green.setFillPattern(CellStyle.SOLID_FOREGROUND);
	style_blue.setFillPattern(CellStyle.SOLID_FOREGROUND);

	String sheetName = "Summary";
	XSSFSheet sheetSummary = wb_out.createSheet(sheetName);
	XSSFRow rowhead = sheetSummary.createRow((short) 0);

	for (int i = 0; i < numberOfCellSummaryRow; i++) {
		XSSFCell cell = rowhead.createCell(i);
		headerFileSummary.add(cell);
	}

	for (int i = 0; i < headerFileSummary.size(); i++) {
		if (i == 0)
			headerFileSummary.get(i).setCellValue("Project Acronym");
		else if (i == 1)
			headerFileSummary.get(i).setCellValue("Project Description");
		else if (i == 2)
			headerFileSummary.get(i).setCellValue("Source System");
		else if (i == 3)
			headerFileSummary.get(i).setCellValue("Source Layer");
		else if (i == 4)
			headerFileSummary.get(i).setCellValue("Source Object");
		else if (i == 5)
			headerFileSummary.get(i).setCellValue("Source Data Format");
		else if (i == 6)
			headerFileSummary.get(i).setCellValue("Target Layer");
		else if (i == 7)
			headerFileSummary.get(i).setCellValue("Target Object");
		else if (i == 8)
			headerFileSummary.get(i).setCellValue("Frequency");
		else if (i == 9)
			headerFileSummary.get(i).setCellValue("Loading Type");
		else if (i == 10)
			headerFileSummary.get(i).setCellValue("Target Key Field");
		else if (i == 11)
			headerFileSummary.get(i).setCellValue("Target Object Description");
	}

	for (int i = 0; i < headerFileSummary.size(); i++) {
		if (i < 6)
			headerFileSummary.get(i).setCellStyle(style_green);
		else if (i >= 6)
			headerFileSummary.get(i).setCellStyle(style_blue);
	}
}

public void fillSummarySheet(XSSFWorkbook wb_out, ArrayList<String> listOfSheet, CellStyles cellStyles) {

	XSSFSheet summarySheet = wb_out.getSheetAt(0);
	XSSFCellStyle defStyle = cellStyles.createDefaultXSSFCellStyle(wb_out);

	for (int i = 0; i < listOfSheet.size(); i++) {
		XSSFRow row = wb_out.getSheet("Summary").createRow((short) i + 1);
		Cell _f = row.createCell(0);
		_f.setCellValue("Java Test");
		_f.setCellStyle(defStyle);
		Cell _f1 = row.createCell(1);
		_f1.setCellValue("blablabla");
		_f1.setCellStyle(defStyle);
		Cell _f2 = row.createCell(2);
		_f2.setCellValue("OSS System");
		_f2.setCellStyle(defStyle);
		Cell _f3 = row.createCell(3);
		_f3.setCellValue("blablabla_SL");
		_f3.setCellStyle(defStyle);
		Cell _f4 = row.createCell(4);
		_f4.setCellValue(listOfSheet.get(i).toLowerCase());
		_f4.setCellStyle(defStyle);
		Cell _f5 = row.createCell(5);
		_f5.setCellValue("Fixed");
		_f5.setCellStyle(defStyle);
		Cell _f6 = row.createCell(6);
		_f6.setCellValue("tl_");
		_f6.setCellStyle(defStyle);
		Cell _f7 = row.createCell(7);
		_f7.setCellValue("ent_" + listOfSheet.get(i).toLowerCase());
		_f7.setCellStyle(defStyle);
		Cell _f8 = row.createCell(8);
		_f8.setCellValue("Daily");
		_f8.setCellStyle(defStyle);
		Cell _f9 = row.createCell(9);
		_f9.setCellValue("Full");
		_f9.setCellStyle(defStyle);
		Cell _f10 = row.createCell(10);
		_f10.setCellValue("key1,key2,key3 ... keyN");
		_f10.setCellStyle(defStyle);
		Cell _f11 = row.createCell(11);
		_f11.setCellValue("BlaBlaBla ");
		_f11.setCellStyle(defStyle);
	}
	int lastRowNum = summarySheet.getLastRowNum();
	summarySheet.removeRow(summarySheet.getRow(1));
	int rowIndex = 1;
	if (rowIndex >= 0 && rowIndex < lastRowNum) {
		summarySheet.shiftRows(rowIndex + 1, lastRowNum, -1);
	}
}

public void autoSizeColumns(XSSFWorkbook workbook) {
	int numberOfSheets = workbook.getNumberOfSheets();
	for (int i = 0; i < numberOfSheets; i++) {
		Sheet sheet = workbook.getSheetAt(i);
		if (sheet.getPhysicalNumberOfRows() > 0) {
			Row row = sheet.getRow(sheet.getFirstRowNum());
			Iterator<Cell> cellIterator = row.cellIterator();
			while (cellIterator.hasNext()) {
				Cell cell = cellIterator.next();
				int columnIndex = cell.getColumnIndex();
				sheet.autoSizeColumn(columnIndex);
			}
		}
	}
}

public void deleteDirtySheet(XSSFWorkbook wb_out) {
	int index = 0;
	if (wb_out.getSheet("ent_table_name") != null) {
		index = wb_out.getSheetIndex(wb_out.getSheet("ent_table_name"));
		wb_out.removeSheetAt(index);
	}
}

@SuppressWarnings("deprecation")
public void generateSheetsHeader(XSSFRow rowhead, XSSFWorkbook wb_out, CellStyles cellStyles) {

	ArrayList<XSSFCell> headerFile = new ArrayList<>();
	XSSFColor green = new XSSFColor(new java.awt.Color(0, 176, 80));
	XSSFColor blue = new XSSFColor(new java.awt.Color(84, 142, 212));
	XSSFCellStyle style_green = cellStyles.createHeaderXSSFCellStyle(wb_out, green);
	XSSFCellStyle style_blue = cellStyles.createHeaderXSSFCellStyle(wb_out, blue);
	style_green.setFillPattern(CellStyle.SOLID_FOREGROUND);
	style_blue.setFillPattern(CellStyle.SOLID_FOREGROUND);

	for (int i = 0; i < numberOfCellDefSheets; i++) {
		XSSFCell cell = rowhead.createCell(i);
		headerFile.add(cell);
	}

	for (int i = 0; i < headerFile.size(); i++) {
		if (i == 0)
			headerFile.get(i).setCellValue("Source System");
		else if (i == 1)
			headerFile.get(i).setCellValue("Source Object");
		else if (i == 2)
			headerFile.get(i).setCellValue("Source Field");
		else if (i == 3)
			headerFile.get(i).setCellValue("Datatype Source Field");
		else if (i == 4)
			headerFile.get(i).setCellValue("Format Source Field");
		else if (i == 5)
			headerFile.get(i).setCellValue("Key Source Field");
		else if (i == 6)
			headerFile.get(i).setCellValue("Target Layer");
		else if (i == 7)
			headerFile.get(i).setCellValue("Target Object");
		else if (i == 8)
			headerFile.get(i).setCellValue("Target Field");
		else if (i == 9)
			headerFile.get(i).setCellValue("Datatype Target Field");
		else if (i == 10)
			headerFile.get(i).setCellValue("Key Target Field");
		else if (i == 11)
			headerFile.get(i).setCellValue("Rules Target Field");
		else if (i == 12)
			headerFile.get(i).setCellValue("Nullable Target Field");
	}

	for (int i = 0; i < headerFile.size(); i++) {
		if (i < 6)
			headerFile.get(i).setCellStyle(style_green);
		else if (i >= 6)
			headerFile.get(i).setCellStyle(style_blue);
	}
}

public String normalizeDataType(String value) {
	if (value.equals("DATE")) {
		value = "date";
	} else if (value.contains("TIMESTAMP")) {
		value = "timestamp";
	} else if (value.contains("VARCHAR")) {
		value = "varchar";
	} else if (value.equals("NUMBER") || value.equals("FLOAT")) {
		value = "numeric";
	} else if (value.equals("CHAR")) {
		value = "varchar";
	} else if (value.equals("CLOB")) {
		value = "text";
	}
	
	return value;
}

public String computeTargetDataValue(ArrayList<String> tempStructure) {
	String targetFiled = "";
	String targetFiledDataType = "";

	String sourceField = "";
	String sourceObj = "";
	String lenghtField = "";
	String dataType = "";
	String nullable = "";

		for (int i = 0; i < tempStructure.size(); i++) {
			if (i == 0)
				sourceObj = tempStructure.get(i).toString();
			else if (i == 1)
				sourceField = tempStructure.get(i).toString();
			else if (i == 2)
				dataType = tempStructure.get(i).toString();
			else if (i == 3)
				lenghtField = tempStructure.get(i).toString();
			else if (i == 4)
				nullable = tempStructure.get(i).toString();
		}
	
	// manage Type -> DATE
		if (dataType.contains("DATE")) {
			if (sourceField.contains("_DATE") || sourceField.contains("DATE")) {
				targetFiled = "dt_" + sourceField.toLowerCase();
				targetFiledDataType = "date";
			} else {
				targetFiled = "dt_" + sourceField.toLowerCase() + "_date";
				targetFiledDataType = "date";
			}
		}
		else if (dataType.contains("TIMESTAMP")) {
			if (sourceField.contains("_TIMESTAMP")) {
				targetFiled = "dt_" + sourceField.toLowerCase();
				targetFiledDataType = "timestamp";
			} else {
				targetFiled = "dt_" + sourceField.toLowerCase() + "_timestamp";
				targetFiledDataType = "timestamp";
			}	
			
		}
	// manage Type -> CODE
		else if (( (!sourceField.contains("_DESC") && (!sourceField.contains("DESC_")))  &&
				   (!sourceField.contains("_ID") && (!sourceField.contains("ID_")))  &&
				   (!sourceField.contains("_FL") && (!sourceField.contains("_FLG"))) && (!sourceField.contains("FLG_") && (!sourceField.contains("FL_"))) && 
				   ( dataType.contains("INT") || dataType.contains("BIGINT") || dataType.contains("NUMBER")	|| 
				   (dataType.contains("CHAR") && Double.parseDouble(lenghtField) > 2 && Double.parseDouble(lenghtField) < 41))) ) {
			
			if ((sourceField.contains("_CODE") || sourceField.contains("CODE") || sourceField.contains("_CD")
					|| sourceField.contains("_COD"))) {
				targetFiled = "cd_" + sourceField.toLowerCase();
				if (dataType.contains("CHAR")) {
					int lenghtFieldInt = (int) Double.parseDouble(lenghtField);
					targetFiledDataType = "varchar(" + lenghtFieldInt + ")";
				} else if (dataType.contains("NUMBER")) {
					String res = lenghtField.replace(".", ",");
					targetFiledDataType = "numeric(" + res + ")";
				} else {
					String value = normalizeDataType(dataType);
					targetFiledDataType = value;
				}
			} 
			else if (sourceField.contains("CODE_") || sourceField.contains("CD_") || sourceField.contains("COD_")
					&& ((!sourceField.contains("_ID") || !sourceField.contains("ID_")))) {
				targetFiled = sourceField.toLowerCase() + "_code";
				if (dataType.contains("CHAR")) {
					int lenghtFieldInt = (int) Double.parseDouble(lenghtField);
					targetFiledDataType = "varchar(" + lenghtFieldInt + ")";
				} else if (dataType.contains("NUMBER")) {
					String res = lenghtField.replace(".", ",");
					targetFiledDataType = "numeric(" + res + ")";
				} else {
					String value = normalizeDataType(dataType);
					targetFiledDataType = value;
				}
			}
			else {
				targetFiled = "cd_" + sourceField.toLowerCase() + "_code";
				if (dataType.contains("VARCHAR")) {
					String value = normalizeDataType(dataType);
					int lenghtFieldInt = (int) Double.parseDouble(lenghtField);
					targetFiledDataType = value + "(" + lenghtFieldInt + ")";
				} else if (dataType.contains("NUMBER")) {
					String res = lenghtField.replace(".", ",");
					targetFiledDataType = "numeric(" + res + ")";
				} else {
					String value = normalizeDataType(dataType);
					targetFiledDataType = value;
				}
			}
		}
	// manage Type -> ID
		else if (dataType.contains("INT") || dataType.contains("BIGINT") || dataType.contains("NUMBER")) {
			
			if (sourceField.contains("_ID")) {
				targetFiled = "id_" + sourceField.toLowerCase();
				String value = normalizeDataType(dataType);
				targetFiledDataType = value;
				if (dataType.contains("NUMBER")) {
					String res = lenghtField.replace(".", ",");
					targetFiledDataType = "numeric(" + res + ")";
				}
			}
			else if (sourceField.contains("ID_")) {
				targetFiled = sourceField.toLowerCase() + "_id";
				String value = normalizeDataType(dataType);
				targetFiledDataType = value;
				if (dataType.contains("NUMBER")) {
					String res = lenghtField.replace(".", ",");
					targetFiledDataType = "numeric(" + res + ")";
				}
			}
			else {
				targetFiled = "id_" + sourceField.toLowerCase() + "_id";
				String value = normalizeDataType(dataType);
				if (dataType.contains("NUMBER")) {
					String res = lenghtField.replace(".", ",");
					targetFiledDataType = "numeric(" + res + ")";
				} else
					targetFiledDataType = value;
			}
		}
	// manage Type -> DESC
		else if (dataType.contains("CLOB") || dataType.contains("TEXT") || dataType.contains("CHAR") && Double.parseDouble(lenghtField) > 1) {
			
			if (sourceField.contains("_DESC")) {
				targetFiled = "ds_" + sourceField.toLowerCase();
				if (Double.parseDouble(lenghtField) > 255) {
					targetFiledDataType = "text";
				} else {
					int lenghtFieldInt = (int) Double.parseDouble(lenghtField);
					targetFiledDataType = "varchar(" + lenghtFieldInt + ")";
				}
			} 
			else if (sourceField.contains("DS_")) {
				targetFiled = sourceField.toLowerCase() + "_desc";
				if (Double.parseDouble(lenghtField) > 255) {
					targetFiledDataType = "text";
				} else {
					int lenghtFieldInt = (int) Double.parseDouble(lenghtField);
					targetFiledDataType = "varchar(" + lenghtFieldInt + ")";
				}
			}
			else if (sourceField.contains("DESCRIPTION")) {
				targetFiled = "ds_" + sourceField.toLowerCase();
				if (Double.parseDouble(lenghtField) > 255) {
					targetFiledDataType = "text";
				} else {
					int lenghtFieldInt = (int) Double.parseDouble(lenghtField);
					targetFiledDataType = "varchar(" + lenghtFieldInt + ")";
				}
			} 
			else
				targetFiled = "ds_" + sourceField.toLowerCase() + "_desc";
			
			if (Double.parseDouble(lenghtField) > 255) {
				targetFiledDataType = "text";
			} 
			else {
				int lenghtFieldInt = (int) Double.parseDouble(lenghtField);
				targetFiledDataType = "varchar(" + lenghtFieldInt + ")";
			}
		}
	// manage Type -> FLAG
		else if (dataType.contains("INT") || (dataType.contains("CHAR") && Double.parseDouble(lenghtField) <= 2) || dataType.contains("BOOLEAN")) {
			
			if (sourceField.contains("FLG_") || sourceField.contains("FL_")) {
				String tmp = sourceField.replaceAll("FLG_", "FL_");
				targetFiled = tmp.toLowerCase() + "_flag";
			} 
			else if (sourceField.contains("_FLAG")) {
				targetFiled = "fl_" + sourceField.toLowerCase();
			} 
			else
				targetFiled = "fl_" + sourceField.toLowerCase() + "_flag";

			if (dataType.contains("CHAR")) {
				int lenghtFieldInt = (int) Double.parseDouble(lenghtField);
				targetFiledDataType = "varchar(" + lenghtFieldInt + ")";
			} 
			else
				targetFiledDataType = lenghtField;
		}
    // manage Type -> NUMBER
		else if (dataType.contains("NUMBER") || dataType.contains("FLOAT")) {
			
			if (sourceField.contains("NM_"))
				targetFiled = sourceField.toLowerCase() + "_number";
			else if (sourceField.contains("_NUMBER")) 
				targetFiled = "nm_" + sourceField.toLowerCase();
		    else
				targetFiled = "nm_" + sourceField.toLowerCase() + "_number";

			if (dataType.contains("NUMBER")) {
				targetFiledDataType = "numeric";
			}
			else if (dataType.contains("FLOAT")) {
				String res = lenghtField.replace(".", ",");
				targetFiledDataType = "numeric(" + res + ")";
			}
		}

		if (!targetFiled.equals("") || !targetFiledDataType.equals(""))
			return sourceObj + "@" + targetFiled + "@" + targetFiledDataType + "@" + nullable;
		else 
			return "";
		
	}

public String manageDataSourceType(ArrayList<String> tempStructure, String dataValue) {
	String concat = "";
	String dataLenght = tempStructure.get(3);
	
	if(dataValue.equals("numeric")) {
		String sub = dataLenght.replace(".", ",");
		concat = dataValue + "("+sub+")"; 
	}
	else if(dataValue.equals("varchar")) {
		int lenghtFieldInt = (int) Double.parseDouble(dataLenght);
		concat = dataValue + "("+lenghtFieldInt+")"; 
	}
	else { 
		concat = dataValue;
	}
	return concat;
}


}