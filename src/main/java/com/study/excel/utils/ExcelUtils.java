package com.study.excel.utils;

import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.study.excel.dto.FieldInfo;

import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.util.*;

/**
 * 
 * @author 050355
 *
 */
public final class ExcelUtils {

	/**
	 * 构造方法
	 */
	private ExcelUtils() {

	}

	private static final Logger logger = LoggerFactory.getLogger(ExcelUtils.class);
	/**
	 * 2007板Excel
	 */
	public static final String EXCEL_FILE_EXTENSION_XLSX = "xlsx";

	/**
	 * 2003 --- 2007板Excel
	 */
	public static final String EXCEL_FILE_EXTENSION_XLS = "xls";

	/**
	 * 默认的时间格试
	 */
	public static final String DEFAULT_DATE_TIME_PATTERN = "yyyy-MM-dd HH:mm:ss";// 默认日期格式

	/**
	 * 默认的时间格试
	 */
	public static final String DEFAULT_DATE_PATTERN = "yyyy-MM-dd";// 默认日期格式

	public static final String DEFAULT_NUMBER_INTEGER_FORMAT = "0";

	public static final String DEFAULT_NUMBER_DOUBLE_FORMAT = "0.00";

	public static final String DEFAULT_NUMBER_BIGDECIMAL_FORMAT = "#,##0.00";

	/**
	 * 默认宽度
	 */
	public static final int DEFAULT_COLUMN_WIDTH = 17;

	/**
	 * 每个sheet页写入的最大行数
	 */

	public static final int DEFALUT_MAX_SHEET_RECORD = 65535;

	/**
	 * 创建Excel 工作薄 如果给定文件名为空， 则返回null;
	 * 
	 * @param fileName
	 * @return
	 * @throws IOException
	 */
	public static Workbook createWorkbook(InputStream input, String fileSurefix) throws IOException {
		Workbook workbook = null;
		if (input != null) {
			if (EXCEL_FILE_EXTENSION_XLSX.equals(fileSurefix)) {
				workbook = new XSSFWorkbook(input);
			} else if (EXCEL_FILE_EXTENSION_XLS.equals(fileSurefix)) {
				workbook = new HSSFWorkbook(input);
			} else {
				workbook = new XSSFWorkbook(input);
			}
		}
		return workbook;

	}

	/**
	 * @return
	 * @throws IOException
	 */
	public static Workbook createXssfWorkbook(InputStream inputStream) throws Exception {
		return new XSSFWorkbook(inputStream);
	}

	/**
	 * 创建单元格样式
	 * 
	 * @param wb
	 * @return
	 */
	public static CellStyle createCellStyle(Workbook wb) {
		return createCommonCellStyle(wb);
	}

	/**
	 * 
	 * 设置 单元格边框为红色
	 * 
	 * @param cellStyle
	 * @return
	 */
	public static void setRedBorderColorStyle(CellStyle cellStyle) {
		setBorderColorStyle(cellStyle, HSSFColor.RED.index);
	}

	/**
	 * 
	 * 设置 单元格边框为红色
	 * 
	 * @param cellStyle
	 * @return
	 */
	public static void setBorderColorStyle(CellStyle cellStyle, short color) {
		cellStyle.setBottomBorderColor(color);
		cellStyle.setTopBorderColor(color);
		cellStyle.setLeftBorderColor(color);
		cellStyle.setRightBorderColor(color);

	}

	/**
	 * 
	 * 设置 单元格边框为红色
	 * 
	 * @param cellStyle
	 * @return
	 */
	public static void setBlackBorderColorStyle(CellStyle cellStyle) {
		setBorderColorStyle(cellStyle, HSSFColor.BLACK.index);
	}

	/**
	 * 
	 * 设置 单元格边框
	 * 
	 * @param cellStyle
	 * @return
	 */
	public static void setBorderThinStyle(CellStyle cellStyle) {
		cellStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN); // 下边框
		cellStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);// 左边框
		cellStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);// 上边框
		cellStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);// 右边框

	}

	/**
	 * 
	 * 设置 单元格边框
	 * 
	 * @param cellStyle
	 * @return
	 */
	public static void setAlignCenter(CellStyle cellStyle) {
		cellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		cellStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER); // 垂直居中

	}

	/**
	 * 
	 * @param cellStyle
	 * @param Font
	 * @param fontSize
	 * @param fontName
	 * @return
	 */

	public static void setFont(CellStyle cellStyle, Font Font, short fontSize, String fontName) {
		Font.setFontHeightInPoints(fontSize); // 字体高度
		Font.setFontName(fontName); // 字体样式
		cellStyle.setFont(Font);

	}

	/**
	 * 
	 * @param cellStyle
	 * @param Font
	 * @param fontSize
	 * @param fontName
	 * @return
	 */

	public static void setTitileFont(CellStyle cellStyle, Font font, short fontSize, String fontName) {
		setFont(cellStyle, font, fontSize, fontName);
		font.setBoldweight(Font.BOLDWEIGHT_BOLD);
	}

	/**
	 * 设置 表头样式
	 * 
	 * @param workbook
	 * @return
	 */
	public static CellStyle createHeaderStyle(Workbook workbook) {
		CellStyle titleStyle = workbook.createCellStyle();
		ExcelUtils.setBorderThinStyle(titleStyle);
		ExcelUtils.setAlignCenter(titleStyle);
		Font titleFont = workbook.createFont();
		ExcelUtils.setTitileFont(titleStyle, titleFont, (short) 20, "黑体");
		return titleStyle;
	}

	/**
	 * 设置 表头样式
	 * 
	 * @param workbook
	 * @return
	 */
	public static CellStyle createTitileStyle(Workbook workbook) {
		CellStyle titleStyle = workbook.createCellStyle();
		ExcelUtils.setBorderThinStyle(titleStyle);
		ExcelUtils.setAlignCenter(titleStyle);
		Font titleFont = workbook.createFont();
		ExcelUtils.setTitileFont(titleStyle, titleFont, (short) 15, "黑体");
		return titleStyle;
	}

	/**
	 * 设置 表头样式
	 * 
	 * @param workbook
	 * @return
	 */
	public static CellStyle createCommonCellStyle(Workbook workbook) {
		CellStyle cellStyle = workbook.createCellStyle();
		ExcelUtils.setBorderThinStyle(cellStyle);
		ExcelUtils.setAlignCenter(cellStyle);
		Font titleFont = workbook.createFont();
		ExcelUtils.setFont(cellStyle, titleFont, (short) 15, "黑体");
		return cellStyle;
	}

	/**
	 * 
	 * @param row
	 * @param cellIndex
	 * @return
	 */

	public static Cell getOrCreateCell(Row row, int cellIndex) {
		Cell cell = row.getCell(cellIndex);
		if (null == cell) {
			cell = row.createCell(cellIndex);
		}
		return cell;
	}

	/**
	 * 获取已有行或创建新行
	 * 
	 * @param sheet
	 * @param rowIndex 列号
	 * @return
	 */
	public static Row getOrCreateRow(Sheet sheet, int rowIndex) {
		Row row = sheet.getRow(rowIndex);
		if (null == row) {
			row = sheet.createRow(rowIndex);
		}
		return row;
	}

	/**
	 * 导出Excel，表头行
	 */
	private static void writeHeaderRow(Sheet sheet, Row titleRow, FieldInfo[] fieldInfoArray, CellStyle headerStyle) {
		int i = 0;
		int columnWidth = 17;
		Cell cell = null;
		for (FieldInfo fieldInfo : fieldInfoArray) {
			cell = titleRow.createCell(i);
			cell.setCellStyle(headerStyle);
			columnWidth =(fieldInfo.getWidth() == null ? 1 : fieldInfo.getWidth());
			cell.setCellValue(fieldInfo.getTitle());
			sheet.setColumnWidth(i++, columnWidth * 256 *17);
		}
	}

	/**
	 * 从给定sheet中读取数据，并返回读取的数据
	 * 
	 * @param <T>      数据类型
	 * @param sheet    Sheet
	 * @param cls      数据类型
	 * @param startRow 起始行
	 * @return
	 * @throws Exception
	 */
	public static <T> List<T> readDateFromSheet(Sheet sheet, Class<T> cls, int startRow) throws Exception {
		List<T> resultData = new ArrayList<T>();
		int lastRowIndex = ExcelUtils.getLastRowIndex(sheet, startRow);
		// int lastRowIndex = sheet.getLastRowNum();
		int i = startRow - 1;
		try {
			FieldInfo[] fieldInfoArr = FieldUtils.getExplortHeadFieldDtos(cls);
			Row row = null;
			while (i < lastRowIndex) {
				T t = FieldUtils.getInstanceByClass(cls);
				row = sheet.getRow(i);
				if (row != null) {
					readerOneRowData(row, fieldInfoArr, t);
					resultData.add(t);
				}
				i++;
			}
		} catch (Exception e) {
			logger.error("读取数所发生异常", e);
			throw new Exception(e);
		}
		return resultData;
	}

	/**
	 * 读取一行数据
	 * 
	 * @param <T>
	 * @param row
	 * @param fieldInfoArr
	 * @param obj
	 * @throws Exception
	 */
	public static <T> void readerOneRowData(Row row, FieldInfo[] fieldInfoArr, T obj) throws Exception {

		int i = 0;
		Cell cell;
		Object value = null;

		for (FieldInfo fieldInfo : fieldInfoArr) {
			cell = row.getCell(i++);
			try {
				value = getObjectValue(cell, fieldInfo);
				fieldInfo.getWriteMethod().invoke(obj, value);
			} catch (Exception e) {
				logger.error("读取" + row.getRowNum() + "行" + fieldInfo.getPropertyName() + "发生错误", e);
				throw new Exception(e);
			}
		}

	}

	/**
	 * 跟据目标属性的类弄，获取单元格中的值
	 * 
	 * @param cell
	 * @param fieldInfo
	 * @return
	 */
	private static Object getObjectValue(Cell cell, FieldInfo fieldInfo) {
		if (cell == null) {
			return null;
		}
		Object value = null;
		Convertor<?> convertor = fieldInfo.getConvertor();
		switch (cell.getCellType()) {
		case XSSFCell.CELL_TYPE_STRING:
			if (convertor != null) {
				value = convertor.inport(cell.getStringCellValue());
			} else {
				value = cell.getStringCellValue();
			}
			/*
			 * if (fieldInfo.getFieldType() == String.class) { // do nothing } else if
			 * (fieldInfo.getFieldType() == Integer.class) { value =
			 * Integer.valueOf((String) value); } else if (fieldInfo.getFieldType() ==
			 * Long.class) { value = Long.valueOf((String) value); } else if
			 * (fieldInfo.getFieldType() == Double.class) { value = Double.valueOf((String)
			 * value); } else if (fieldInfo.getFieldType() == BigDecimal.class) { value =
			 * new BigDecimal((String) value); } else if (fieldInfo.getFieldType() ==
			 * Float.class) { value = Float.valueOf((String) value); } else if
			 * (fieldInfo.getFieldType() == Boolean.class) { value =
			 * Boolean.valueOf((String) value); } else { // do nothing }
			 */
			break;
		case XSSFCell.CELL_TYPE_NUMERIC:
			if (HSSFDateUtil.isCellDateFormatted(cell)) {
				if (convertor != null) {
					value = convertor.inport(cell.getDateCellValue());
				} else {
					value = cell.getDateCellValue();
					if (fieldInfo.getFieldType() == String.class) {
						value = fieldInfo.getDataFormater().format((Date) value);
					}
				}

			} else {

				if (convertor != null) {
					value = convertor.inport(cell.getNumericCellValue());
				} else {
					value = cell.getNumericCellValue();
					if (fieldInfo.getFieldType() == Integer.class) {
						value = (Integer) value;
					} else if (fieldInfo.getFieldType() == Long.class) {
						value = (Long) value;
					} else if (fieldInfo.getFieldType() == BigDecimal.class) {
						value = new BigDecimal((Double) value);
					} else if (fieldInfo.getFieldType() == Float.class) {
						value = (Float) value;
					} else {
						// do nothing
					}

				}
			}
			break;
		case XSSFCell.CELL_TYPE_FORMULA:
			if (convertor != null) {
				value = convertor.inport(cell.getCellFormula());
			} else {
				value = cell.getCellFormula();
			}

			break;
		case XSSFCell.CELL_TYPE_BOOLEAN:
			if (convertor != null) {
				value = convertor.inport(cell.getBooleanCellValue());
			} else {
				value = cell.getBooleanCellValue();
				if (fieldInfo.getFieldType() == String.class) {
					value = String.valueOf(value);
				}
			}
			break;
		default:
			break;
		}
		return value;

	}

	/**
	 * 获取最后一行数据的行索引
	 * 
	 * @param sheet
	 * @param rowStartIndex
	 * @return
	 */
	private static int getLastRowIndex(Sheet sheet, int rowStartIndex) {

		int i = rowStartIndex;

		Row row = sheet.getRow(i);
		Cell cell = row.getCell(0);
		while (row != null && cell != null && getCellValue(cell) != null) {
			i++;
			row = sheet.getRow(i);
			if (!Objects.isNull(row)) {
				cell = row.getCell(0);
			}
		}
		return i;
	}

	/**
	 * 将数据写入到Excel 当中
	 * 
	 * @param <T>
	 * @param book
	 * @param sheet
	 * @param listData
	 * @param startRow
	 * @param writeTitle
	 * @throws Exception
	 */
	public static <T> void writeDateToSheet(Workbook book, Sheet sheet, List<T> listData, int startRow,
			boolean writeTitle) throws Exception {

		if (CollectionUtils.isEmpty(listData)) {
			return;
		}
		Class<?> cls = listData.get(0).getClass();
		try {
			FieldInfo[] fieldInfoArr = FieldUtils.getExplortHeadFieldDtos(cls);
			if (writeTitle) {
				Row titleRow = sheet.createRow(startRow);
				writeHeaderRow(sheet, titleRow, fieldInfoArr, ExcelUtils.createHeaderStyle(book));
				startRow++;
			}
			Row row = null;
			for (T t : listData) {
				row = sheet.createRow(startRow);
				writeOneData(row, t, fieldInfoArr, ExcelUtils.createCommonCellStyle(book));
				startRow++;
			}
		} catch (Exception e) {
			logger.error("导出数据到excel失败！", e);
			throw new Exception(e);
		}

	}

	/**
	 * 写一行数据到excel 行
	 * 
	 * @param <T>
	 * @param row
	 * @param t
	 * @param fieldInfoArr
	 * @param cellStyle
	 * @throws Exception
	 */
	private static <T> void writeOneData(Row row, T t, FieldInfo[] fieldInfoArr, CellStyle cellStyle) throws Exception {
		int len = fieldInfoArr.length;
		Object value = null;
		Cell cell = null;
		FieldInfo fieldInfo = null;
		try {
			for (int i = 0; i < len; i++) {
				cell = row.createCell(i);
				fieldInfo = fieldInfoArr[i];
				value = fieldInfo.getReaderMethod().invoke(t);
				writeDataToCell(cell, value, fieldInfo, cellStyle);
			}
		} catch (Exception e) {
			logger.error("写数据到Excel第" + row.getRowNum() + "出错！", e);
			throw new Exception(e);
		}

	}

	/**
	 * 从sheet里面读取数据
	 * 
	 * @param mapIndex
	 * @param titleRowNum
	 * @param lastRow
	 * @param list
	 * @param sheet
	 */
	private static void writeDataToCell(Cell cell, Object value, FieldInfo fileInfo, CellStyle cellStyle) {
		if (cell == null) {
			return;
		}
		if (value == null) {
			return;
		}
		if (value instanceof Date) {
			cell.setCellStyle(cellStyle);
			if (fileInfo != null && fileInfo.getDateFormat() != null) {
				cellStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat(fileInfo.getDateFormat()));
			} else {
				cellStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat(DEFAULT_DATE_TIME_PATTERN));
			}
			cell.setCellValue((Date) value);
		} else if (value instanceof Number) {
			if (value instanceof Integer || value instanceof Short) {
				cellStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat(DEFAULT_NUMBER_INTEGER_FORMAT));
				cell.setCellStyle(cellStyle);
				cell.setCellValue(Double.valueOf(String.valueOf(value)));
			} else if (value instanceof Float || value instanceof Double) {
				cellStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat(DEFAULT_NUMBER_DOUBLE_FORMAT));
				cell.setCellStyle(cellStyle);
				cell.setCellValue((Double) value);
			} else if (value instanceof BigDecimal) {// BigDecimal默认表示为金额，
				cellStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat(DEFAULT_NUMBER_BIGDECIMAL_FORMAT));
				cell.setCellStyle(cellStyle);
				cell.setCellValue(((BigDecimal) value).doubleValue());
			} else {
				cell.setCellStyle(cellStyle);
				cell.setCellValue(Double.valueOf(String.valueOf(value)));
			}
		} else if (value instanceof Boolean) {
			cell.setCellValue((Boolean) value);
			cell.setCellStyle(cellStyle);
		} else {
			cell.setCellValue(String.valueOf(value));
			cell.setCellStyle(cellStyle);
		}
	}

	/**
	 * 获取单元格的值
	 * 
	 * @param cell
	 * @return
	 */
	public static Object getCellValue(Cell cell) {
		Object value = null;
		if (cell != null) {
			switch (cell.getCellType()) {
			case XSSFCell.CELL_TYPE_STRING:
				value = cell.getStringCellValue();
				break;
			case XSSFCell.CELL_TYPE_NUMERIC:
				String cellValue = null;
				if (DateUtil.isCellDateFormatted(cell)) {
					value = cell.getDateCellValue();
				} else {
					value = cell.getNumericCellValue();
				}
				value = cellValue;
				break;
			case XSSFCell.CELL_TYPE_FORMULA:
				value = cell.getCellFormula();
				break;
			case XSSFCell.CELL_TYPE_BOOLEAN:
				value = cell.getBooleanCellValue();
				break;
			default:
				break;
			}
			return value;
		}
		return null;
	}

	/**
	 * 将某个单元格的边框设置为黑色
	 * 
	 * @param workBook    工作薄
	 * @param recordCount 行最大值
	 * @param cellSize    单元格行结束最大值
	 */
	public static void setBlackColorStyle(Workbook workBook, int recordCount, int cellSize) {
		if (recordCount == 0) {
			return;
		}
		Sheet sheet = workBook.getSheetAt(0);
		recordCount = recordCount + 2;
		Row row = null;
		Cell cell = null;
		CellStyle cellStyle = workBook.createCellStyle();
		for (int i = 2; i < recordCount; i++) {
			row = sheet.getRow(i);
			if (row != null) {
				for (int j = 0; j < cellSize; j++) {
					cell = row.getCell(j);
					if (cell != null) {
						ExcelUtils.setBlackBorderColorStyle(cellStyle);
						cell.setCellStyle(cellStyle);
					}
				}
			}
		}
	}

	/**
	 * 将指定单元格的边框设置为红色
	 * 
	 * @param workBook         工作薄
	 * @param invalidCellIndex 单元格行结束最大值
	 */
	public static void setRedColorStyle(Workbook workBook, Map<Integer, Set<Integer>> invalidCellIndex) {
		if (invalidCellIndex == null || invalidCellIndex.isEmpty()) {
			return;
		}
		Sheet sheet = workBook.getSheetAt(0);
		Set<Integer> keySet = invalidCellIndex.keySet();
		Row row = null;
		Cell cell = null;
		CellStyle cellStyle = workBook.createCellStyle();
		Set<Integer> cellIndexSet = null;
		for (Integer rowIndex : keySet) {
			row = sheet.getRow(rowIndex);
			if (row != null) {
				cellIndexSet = invalidCellIndex.get(rowIndex);
				if (cellIndexSet != null && !cellIndexSet.isEmpty()) {
					for (Integer cellIndex : cellIndexSet) {
						cell = row.getCell(cellIndex);
						if (cell == null) {
							cell = row.createCell(cellIndex);
						}
						if (cell != null) {
							ExcelUtils.setBorderThinStyle(cellStyle);
							ExcelUtils.setRedBorderColorStyle(cellStyle);
							cell.setCellStyle(cellStyle);
						}
					}
				}
			}
		}

	}

}
