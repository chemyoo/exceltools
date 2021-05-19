package com.pers.chemyoo;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.io.FileUtils;
import org.apache.commons.io.IOUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;

import net.sf.jxls.exception.ParsePropertyException;
import net.sf.jxls.transformer.XLSTransformer;

public class ExportExcelUtils
{
	private ExportExcelUtils()
	{
		throw new AbstractMethodError("ExportExcelUtils can not be instanced.");
	}

	/**
	 * 根据模版导出Excel。可导出单页签和多页签。</br>
	 * 多页签时使用Map的key做区分，例如 map.put(sheet1, datas1), map.put(sheet2, datas2);
	 * 
	 * @param is
	 * @param out
	 * @param beanParams
	 */
	@SuppressWarnings("deprecation")
	public static void exportSingleSheet(InputStream is, OutputStream out, Map<String, Object> beanParams)
	{
		try
		{
			XLSTransformer transformer = new XLSTransformer();
			Workbook workbook = transformer.transformXLS(is, beanParams);
			workbook.write(out);
		}
		catch (ParsePropertyException | InvalidFormatException | IOException e)
		{
			LoggerUtils.error(ExportExcelUtils.class, e.getMessage(), e);
		}
		finally
		{
			IOUtils.closeQuietly(out);
		}
	}

	public static void main(String[] args) throws IOException
	{
		InputStream is = FileUtils.openInputStream(new File("D:/2021.xlsx"));
		OutputStream out = new FileOutputStream("D:/123456.xlsx");
		Map<String, Object> beanParams = new HashMap<>();
		List<Map<String, Object>> dataList = new ArrayList<>();
		List<Map<String, Object>> dataList1 = new ArrayList<>();
		for (int i = 0; i < 30; i++)
		{
			Map<String, Object> params = new HashMap<>();
			params.put("order", i + 1);
			params.put("name", "chemyoo" + i);
			params.put("desc", "描述" + i);
			dataList.add(params);
			if (i < 10)
			{
				dataList1.add(params);
			}
		}
		beanParams.put("dataList", dataList);
		beanParams.put("dataList1", dataList1);
		ExportExcelUtils.exportSingleSheet(is, out, beanParams);
	}

}
