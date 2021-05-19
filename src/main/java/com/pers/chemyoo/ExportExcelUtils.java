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

//	XLSTransformer transformer = new XLSTransformer(); 
//    File template = ResourceUtils.getFile("classpath:template/excel/claim_summary_report.xls");
//    InputStream is = new FileInputStream(template); 
//    Workbook workbook = transformer.transformMultipleSheetsList(is,results,monthNames, "results",new HashMap(),1); 
//
//
//    public Workbook transformMultipleSheetsList(InputStream is, List objects, List newSheetNames, String beanName, Map beanParams, int startSheetNum) throws ParsePropertyException
//    该方法里面的参数说明如下：
//    1）is：即Template文件的一个输入流
//    2）newSheetNames：即形成Excel文件的时候Sheet的Name
//    3）objects：即我们传入的对应每个Sheet的一个Java对象，这里传入的List的元素为一个Map对象
//    4）beanName：这个参数在jxls对我们传入的List进行解析的时候使用，而且，该参数还对应Template文件中的Tag，例如，beanName为map，那么在Template文件中取值的公式应该定义成${map.get("property1")}；如果beanName为payslip，公式应该定义成${payslip.get("property1")}
//    5）beanParams：这个参数在使用的时候我的代码没有使用到，这个参数是在如果传入的objects还与其他的对象关联的时候使用的，该参数是一个HashMap类型的参数，如果不使用的话，直接传入new HashMap()即可
//    6）startSheetNo：传入0即可，即SheetNo从0开始

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
