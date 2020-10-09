package com.pers.chemyoo;

import org.slf4j.LoggerFactory;

/**
 * @author liujianqing 2020年4月24日 上午10:22:26
 */
public class LoggerUtils
{

	private LoggerUtils() throws NoSuchMethodException
	{
		throw new NoSuchMethodException("LoggerUtils can not be instansed");
	}

	public static void info(Class<?> clazz, String format, Object... args)
	{
		LoggerFactory.getLogger(clazz).info(format, args);
	}

	public static void error(Class<?> clazz, String msg, Throwable t)
	{
		LoggerFactory.getLogger(clazz).error(msg, t);
	}

	public static void error(Class<?> clazz, String msg)
	{
		LoggerFactory.getLogger(clazz).error(msg);
	}

}
