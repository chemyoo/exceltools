package com.pers.chemyoo;

import java.util.List;
import java.util.Map;

import com.google.common.collect.Lists;

/**
 * 单行表头Excel文件数据读取
 * 
 * @author jianqing.liu
 * @since 2020年9月4日 下午3:49:57
 */
public class ExcelModel
{
	private List<String> heads = null;

	private List<Map<String, Object>> bodys = null;

	private String templateName;

	public List<String> getHeads()
	{
		return heads == null ? Lists.newArrayList() : heads;
	}

	public void setHeads(List<String> heads)
	{
		if (heads == null)
		{
			this.heads = Lists.newArrayList();
		}
		else
		{
			this.heads = heads;
		}
	}

	public List<Map<String, Object>> getBodys()
	{
		return bodys;
	}

	public void setBodys(List<Map<String, Object>> bodys)
	{
		if (bodys == null)
		{
			this.bodys = Lists.newArrayList();
		}
		else
		{
			this.bodys = bodys;
		}
	}

	public String getTemplateName()
	{
		return templateName;
	}

	public void setTemplateName(String templateName)
	{
		this.templateName = templateName;
	}

	@Override
	public String toString()
	{
		return String.format("ExcelModel [heads=%s, bodys=%s, templateName=%s]", heads, bodys, templateName);
	}
}
