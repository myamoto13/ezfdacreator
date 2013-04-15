package com.extia.fdaprocessor.io;

import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class ExcelIO {
	
	public boolean isValid(String... excelFieldList) {
		boolean result = false;
		if (excelFieldList != null) {
			for (int i = 0; i < excelFieldList.length; i++) {
				if (excelFieldList[i] != null && !"".equals(excelFieldList[i])) {
					result = true;
					break;
				}
			}
		}
		return result;
	}
	
	public void setCellValue(String value, int rowIndex, int colIndex, Sheet sheet, boolean createCell) {
		Cell cell = getCell(rowIndex, colIndex, sheet);
		
		if(cell == null && createCell){
			Row row = sheet.getRow(rowIndex);
			if(row == null){
				row = sheet.createRow(rowIndex);
			}
			cell = row.getCell(colIndex);
			if(cell == null){
				cell = row.createCell(colIndex);
			}
		}
		if (cell != null) {
			cell.setCellValue(value);
		}
	}

	public void setCellValue(String value, int rowIndex, int colIndex, Sheet sheet) {
		setCellValue(value, rowIndex, colIndex, sheet, false);
	}

	public Cell getCell(int rowIndex, int colIndex, Sheet sheet) {
		Cell result = null;
		if (sheet != null) {
			Row row = sheet.getRow(rowIndex);
			if (row != null) {
				result = row.getCell(colIndex);
			}
		}
		return result;
	}

	public String getCellValue(int rowIndex, int colIndex, Sheet sheet) {
		return getCellValue(getCell(rowIndex, colIndex, sheet));
	}

	public String getCellValue(Cell cell) {
		String result = null;
		if (cell != null) {
			// TODO : check for nullpointers
			FormulaEvaluator formulaEv = cell.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator();

			CellValue cValue = formulaEv.evaluate(cell);
			switch (cell.getCellType()) {
			case Cell.CELL_TYPE_BOOLEAN:
				result = "" + cell.getBooleanCellValue();
				break;
			case Cell.CELL_TYPE_NUMERIC:
				if (HSSFDateUtil.isCellDateFormatted(cell)) {
					Date date = HSSFDateUtil.getJavaDate(cValue.getNumberValue());
					result = new SimpleDateFormat("dd/mm/YYYY").format(date);
				} else {
					result = "" + cell.getNumericCellValue();
				}
				break;
			case Cell.CELL_TYPE_ERROR:
				result = "" + cell.getErrorCellValue();
				break;
			case Cell.CELL_TYPE_FORMULA:
				switch(cell.getCachedFormulaResultType()) {
					case Cell.CELL_TYPE_NUMERIC:
						result = "" + cell.getNumericCellValue();
						break;
					case Cell.CELL_TYPE_STRING:
						result = "" + cell.getStringCellValue();
						break;
				}
				break;
			case Cell.CELL_TYPE_BLANK:
			case Cell.CELL_TYPE_STRING:
				result = cell.getStringCellValue();
				break;
			}
		}
		return result;
	}
}
