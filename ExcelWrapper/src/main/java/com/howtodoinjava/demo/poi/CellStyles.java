package com.howtodoinjava.demo.poi;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CellStyles {

	public CellStyles() {
	};

	@SuppressWarnings("deprecation")
	public XSSFCellStyle createHeaderXSSFCellStyle(XSSFWorkbook wb, XSSFColor color) {

		XSSFCellStyle cellStyle = wb.createCellStyle();
		XSSFFont headerFont = wb.createFont();
		headerFont.setFontHeightInPoints((short) 11);
		headerFont.setFontName("Calibri");
		headerFont.setColor(IndexedColors.WHITE.getIndex());
		headerFont.setBold(true);
		headerFont.setItalic(false);

		cellStyle.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		cellStyle.setBorderBottom(BorderStyle.THIN);
		cellStyle.setBorderLeft(BorderStyle.THIN);
		cellStyle.setBorderTop(BorderStyle.THIN);
		cellStyle.setBorderRight(BorderStyle.THIN);
		cellStyle.setFillForegroundColor(color);
		cellStyle.setFont(headerFont);
		return cellStyle;
	}

	public XSSFCellStyle createDefaultXSSFCellStyle(XSSFWorkbook wb) {

		XSSFCellStyle cellStyle = wb.createCellStyle();
		XSSFFont cellFont = wb.createFont();
		cellFont.setFontHeightInPoints((short) 11);
		cellFont.setFontName("Calibri");
		cellFont.setColor(IndexedColors.BLACK.getIndex());
		cellFont.setBold(false);
		cellFont.setItalic(false);

		cellStyle.setBorderBottom(BorderStyle.THIN);
		cellStyle.setBorderLeft(BorderStyle.THIN);
		cellStyle.setBorderTop(BorderStyle.THIN);
		cellStyle.setBorderRight(BorderStyle.THIN);
		cellStyle.setFont(cellFont);
		return cellStyle;
	}

}
