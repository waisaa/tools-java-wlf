package cn.wlftool.excel.entity;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;

public class ExcelTest {

	public static void main(String[] args) {
		String filepath = "model/check/pollutant-import.xlsx";
		String validErrorPromptFilepath = "model/check/error-prompt.xlsx";
		try (
				FileInputStream fis = new FileInputStream(filepath);
				XSSFWorkbook workbook = new XSSFWorkbook(fis);
				FileOutputStream fos = new FileOutputStream(validErrorPromptFilepath);
		) {
			XSSFSheet sheet = workbook.getSheetAt(0);
			String promptTitle = "错误提示";
			String promptContent = "该单元格的值应为数值";
			int rowIndex = 50, colIndex = 1;
			XSSFRow row = sheet.getRow(rowIndex);
			if (null == row) {
				row = sheet.createRow(rowIndex);
			}
			XSSFCell cell = row.getCell(colIndex);
			if (null == cell) {
				cell = row.createCell(colIndex);
			}
			setCellBgColor(workbook, cell, IndexedColors.RED);
			setCellPrompt(sheet, promptTitle, promptContent, rowIndex, colIndex);
			sheet.protectSheet("123456");
			workbook.write(fos);
		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	/**
	 * 设置单元格背景
	 */
	public static void setCellBgColor(XSSFWorkbook workbook, XSSFCell cell, IndexedColors color) {
		XSSFCellStyle cellStyle = getDefaultCellStyle(workbook, color, false, true);
		cell.setCellStyle(cellStyle);
	}

	/**
	 * 设置单元格上提示
	 */
	public static void setCellPrompt(XSSFSheet sheet, String promptTitle, String promptContent, int rowIndex, int colIndex) {
		XSSFDataValidationHelper helper = new XSSFDataValidationHelper(sheet);
		DataValidationConstraint constraint = helper.createCustomConstraint("A0");
		CellRangeAddressList region = new CellRangeAddressList(rowIndex, rowIndex, colIndex, colIndex);
		DataValidation validation = helper.createValidation(constraint, region);
		validation.createPromptBox(promptTitle, promptContent);
		validation.setShowPromptBox(true);
		sheet.addValidationData(validation);
	}

	/**
	 * 获取默认格式的单元格样式
	 */
	private static XSSFCellStyle getDefaultCellStyle(XSSFWorkbook workbook, IndexedColors color, boolean bold, boolean lock) {
		XSSFCellStyle cellStyle = workbook.createCellStyle();
		cellStyle.setAlignment(HorizontalAlignment.CENTER);
		cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
		XSSFFont font = workbook.createFont();
		font.setFontName(Const.FONT_TNR);
		font.setBold(bold);
		cellStyle.setFont(font);
		XSSFDataFormat dataFormat = workbook.createDataFormat();
		cellStyle.setDataFormat(dataFormat.getFormat(Const.FMT_TEXT));
		cellStyle.setBorderBottom(BorderStyle.THIN);
		cellStyle.setBorderTop(BorderStyle.THIN);
		cellStyle.setBorderLeft(BorderStyle.THIN);
		cellStyle.setBorderRight(BorderStyle.THIN);
		if (null != color) {
			cellStyle.setFillForegroundColor(color.getIndex());
			cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		}
		cellStyle.setLocked(lock);
		return cellStyle;
	}
}
