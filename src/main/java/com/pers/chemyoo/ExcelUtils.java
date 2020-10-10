package com.pers.chemyoo;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.io.FileUtils;
import org.apache.commons.io.IOUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.Validate;
import org.apache.commons.lang3.time.FastDateFormat;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;

import com.google.common.collect.Lists;
import com.google.common.collect.Maps;

/**
 * 读取Excel文件，支持多表头读取
 * 
 * @author chemyoo
 */
public class ExcelUtils
{

	private static FastDateFormat dateFormat = FastDateFormat.getInstance("yyyy-MM-dd");

	private ExcelUtils()
	{
		throw new AbstractMethodError("ExcelUtils can not be instanced.");
	}

	public static final String OFFICE_EXCEL_2003_POSTFIX = ".xls";
	public static final String OFFICE_EXCEL_2010_POSTFIX = ".xlsx";

	public static final String EMPTY = "";

	/**
	 * 读取Excel转成WorkBook
	 * 
	 * @param file
	 * @return
	 * @throws IOException
	 * @throws InvalidFormatException
	 * @throws EncryptedDocumentException
	 */
	public static Workbook readExcel(InputStream is) throws InvalidFormatException, IOException
	{
		return WorkbookFactory.create(is);
	}

	private static ExcelModel read2ExcelModel(Sheet sheet, ExcelConfig config)
	{
		ExcelModel model = new ExcelModel();
		if (sheet == null)
		{
			return model;
		}
		int end = sheet.getLastRowNum();
		int start = config.getStartRow() - 1;
		if (start < 0 || start > end)
		{
			start = 0;
		}
		List<Map<String, Object>> datas = new ArrayList<>();
		if (config.isHasHead())
		{
			model.setHeads(getHeads(sheet, config));
		}
		int headLength = model.getHeads().size();
		for (; start <= end; start++)
		{
			Row row = sheet.getRow(start);
			if (row == null)
			{
				row = sheet.createRow(start);
			}
			int rows = row.getLastCellNum();
			Map<String, Object> map = new HashMap<>();
			for (int j = 0; j < rows; j++)
			{
				String title = Integer.toString(j);
				if (j < headLength)
				{
					title = model.getHeads().get(j);
				}
				String value = getCellString(row.getCell(j));
				if (!EMPTY.equals(value))
				{
					map.put(title, value);
				}
			}
			// 过滤空行
			addRowData(datas, map);
		}
		model.setBodys(datas);
		return model;
	}

	public static Map<String, ExcelModel> read2ExcelModel(InputStream is, ExcelConfig... configs)
	{
		Validate.notNull(is, "读取文件失败");
		Map<String, ExcelModel> map = Maps.newHashMap();
		try
		{
			Workbook book = WorkbookFactory.create(is);
			for (ExcelConfig config : configs)
			{
				int sheetNumber = book.getNumberOfSheets();
				Sheet sheet = book.getSheetAt(config.getSheetIndex() >= sheetNumber ? sheetNumber - 1 : config.getSheetIndex());
				if (!sheet.getSheetName().equalsIgnoreCase(config.getSheetName()))
				{
					sheet = book.getSheet(config.getSheetName());
				}
				map.put(config.getAliasName(), read2ExcelModel(sheet, config));
			}
		}
		catch (InvalidFormatException | IOException e)
		{
			LoggerUtils.error(ExcelUtils.class, e.getMessage(), e);
		}
		finally
		{
			IOUtils.closeQuietly(is);
		}
		return map;
	}

	private static ExcelModel read2ExcelModel(InputStream is, ExcelConfig config)
	{
		Validate.notNull(is, "读取文件失败");
		ExcelModel model = new ExcelModel();
		try
		{
			Workbook book = WorkbookFactory.create(is);
			Sheet sheet = book.getSheetAt(config.getSheetIndex());
			if (!sheet.getSheetName().equalsIgnoreCase(config.getSheetName()))
			{
				sheet = book.getSheet(config.getSheetName());
			}
			return read2ExcelModel(sheet, config);
		}
		catch (InvalidFormatException | IOException e)
		{
			LoggerUtils.error(ExcelUtils.class, e.getMessage(), e);
		}
		finally
		{
			IOUtils.closeQuietly(is);
		}
		return model;
	}

	/**
	 * 表头数据读取
	 * 
	 * @param is 文件流
	 * @param hasHead 是否有表头
	 * @param startRow 数据开始行，从1开始
	 * @return
	 */
	public static ExcelModel read2ExcelModel(InputStream is, boolean hasHead, int startRow)
	{
		ExcelConfig config = new ExcelConfig();
		config.setStartRow(startRow);
		config.setHasHead(hasHead);
		return read2ExcelModel(is, config);
	}

	private static void addRowData(List<Map<String, Object>> datas, Map<String, Object> map)
	{
		if (!map.isEmpty())
		{
			datas.add(map);
		}
	}

	private static List<String> getHeads(Sheet sheet, ExcelConfig config)
	{
		List<String> heads = Lists.newArrayList();
		// 单行表头
		if (config.getHeadStart() == config.getHeadEnd())
		{
			Row row = sheet.getRow(config.getHeadStart() - 1);
			int cells = row.getLastCellNum();
			for (int j = 0; j < cells; j++)
			{
				Cell cell = row.getCell(j);
				if (cell != null && StringUtils.isNotBlank(getCellString(cell)))
				{
					heads.add(getCellString(cell));
				}
			}
		}
		else if (config.getHeadEnd() - config.getHeadStart() > 0)
		{
			// 多行表头
			int rows = config.getHeadEnd() - config.getHeadStart();
			ExcelUtils.fillMerged(sheet, config);
			List<Map<Integer, String>> headRows = Lists.newArrayListWithExpectedSize(rows + 1);
			for (int i = config.getHeadStart() - 1; i < config.getHeadEnd(); i++)
			{
				Row row = sheet.getRow(i);
				int cells = row.getLastCellNum();
				Map<Integer, String> titles = Maps.newHashMapWithExpectedSize(cells);
				for (int j = 0; j < cells; j++)
				{
					Cell cell = row.getCell(j);
					if (cell != null && StringUtils.isNotBlank(getCellString(cell)))
					{
						titles.put(j, getCellString(cell));
					}
				}
				headRows.add(titles);
			}
			heads.addAll(ExcelUtils.mergedHeads(headRows));
		}
		return heads;
	}

	/**
	 * 合并表头
	 * 
	 * @param headRows
	 * @return
	 */
	private static List<String> mergedHeads(List<Map<Integer, String>> headRows)
	{
		if (!headRows.isEmpty())
		{
			int size = headRows.get(0).size();
			List<String> heads = Lists.newArrayListWithCapacity(size);
			for (int i = 0; i <= size; i++)
			{
				StringBuilder titleBuider = new StringBuilder();
				for (Map<Integer, String> hr : headRows)
				{
					String v = hr.get(i);
					if (v != null && titleBuider.length() == 0)
					{
						titleBuider.append(v);
					}
					else if (v != null && !titleBuider.toString().endsWith(v))
					{
						titleBuider.append(".").append(v);
					}
				}
				heads.add(titleBuider.toString());
			}
			return heads;
		}
		return Lists.newArrayList();
	}

	/**
	 * 填充合并单元格的值
	 * 
	 * @param sheet
	 * @return
	 */
	private static void fillMerged(Sheet sheet, ExcelConfig config)
	{
		int sheetMergeCount = sheet.getNumMergedRegions();
		for (int i = 0; i < sheetMergeCount; i++)
		{
			CellRangeAddress range = sheet.getMergedRegion(i);
			int firstColumn = range.getFirstColumn();
			int lastColumn = range.getLastColumn();
			int firstRow = range.getFirstRow();
			int lastRow = range.getLastRow();
			if (firstRow >= config.getHeadStart() - 1 && lastRow < config.getHeadEnd())
			{
				// 合并单元格值填充
				Row row = sheet.getRow(firstRow);
				Cell cell = row.getCell(firstColumn);
				String value = getCellString(cell);
				for (int j = firstRow; j <= lastRow; j++)
				{
					for (int n = firstColumn; n <= lastColumn; n++)
					{
						cell = ExcelUtils.getCell(sheet.getRow(j), n);
						cell.setCellValue(value);
					}
				}
			}
		}
	}

	private static Cell getCell(Row row, int colunm)
	{
		Cell cell = row.getCell(colunm);
		if (cell == null)
		{
			cell = row.createCell(colunm);
		}
		return cell;
	}

	private static String getCellString(Cell cell)
	{
		if (cell == null)
		{
			return EMPTY;
		}
		DataFormatter dataFormatter = new DataFormatter();
		String cellValue = null;
		if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC)
		{
			// cellValue = HSSFDateUtil.getJavaDate(cell.getNumericCellValue());
			if (HSSFDateUtil.isCellDateFormatted(cell))
			{
				cellValue = dateFormat.format(cell.getDateCellValue()); // 日期型
			}
			else
			{
				cellValue = dataFormatter.formatCellValue(cell);
//				String.valueOf(cell.getNumericCellValue()); // 数字
			}
		}
		else if (cell.getCellType() == Cell.CELL_TYPE_FORMULA)
		{
			cell.setCellType(Cell.CELL_TYPE_STRING);
			cellValue = String.valueOf(cell.getStringCellValue());
		}
		else
		{
			cellValue = String.valueOf(cell.getStringCellValue());
		}
		return cellValue;
	}
	
	public static void main(String[] args) throws IOException
	{
		File file = com.pers.chemyoo.exceltools.ExcelUtils.getFile();
		ExcelConfig[] configs = new ExcelConfig[2];
		ExcelConfig config = new ExcelConfig();
		config.setAliasName("哈哈哈哈");
		config.setHasHead(true);
		config.setSheetIndex(6);
		config.setSheetName("变电站建筑工程费用汇总表");
		config.setStartRow(5);
		config.setHeadStart(3);
		config.setHeadEnd(4);
		configs[0] = config;
		config = new ExcelConfig();
		config.setHasHead(true);
		config.setSheetIndex(7);
		config.setSheetName("建筑分部分项工程量清单计价表");
		config.setStartRow(7);
		config.setHeadStart(3);
		config.setHeadEnd(6);
		configs[1] = config;
		System.err.println(ExcelUtils.read2ExcelModel(FileUtils.openInputStream(file), configs));
	}

}
