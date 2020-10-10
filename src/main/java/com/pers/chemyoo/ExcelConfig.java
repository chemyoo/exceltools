package com.pers.chemyoo;

import org.apache.commons.lang3.StringUtils;

public class ExcelConfig
{
	private String sheetName;

	/**
	 * sheet页签别名
	 */
	private String aliasName;

	private int startRow;

	private int sheetIndex = 0;

	private boolean hasHead;

	private int headStart;

	private int headEnd;

	private boolean sheetNameKey = false;

	public String getSheetName()
	{
		return sheetName;
	}

	public void setSheetName(String sheetName)
	{
		this.sheetName = sheetName;
	}

	public int getStartRow()
	{
		return startRow;
	}

	public void setStartRow(int startRow)
	{
		this.startRow = startRow;
	}

	public boolean isHasHead()
	{
		return hasHead;
	}

	public void setHasHead(boolean hasHead)
	{
		this.hasHead = hasHead;
	}

	public int getSheetIndex()
	{
		return sheetIndex;
	}

	public void setSheetIndex(int sheetIndex)
	{
		this.sheetIndex = sheetIndex;
	}

	public boolean isSheetNameKey()
	{
		return sheetNameKey;
	}

	public void setSheetNameKey(boolean sheetNameKey)
	{
		this.sheetNameKey = sheetNameKey;
	}

	public int getHeadStart()
	{
		return headStart;
	}

	public void setHeadStart(int headStart)
	{
		this.headStart = headStart;
	}

	public int getHeadEnd()
	{
		return headEnd;
	}

	public void setHeadEnd(int headEnd)
	{
		this.headEnd = headEnd;
	}

	/**
	 * 如果aliasName为空，则取sheetName。
	 * @return
	 */
	public String getAliasName()
	{
		if (StringUtils.isNotBlank(aliasName))
		{
			return aliasName;
		}
		return sheetName;
	}

	public void setAliasName(String aliasName)
	{
		this.aliasName = aliasName;
	}

}
