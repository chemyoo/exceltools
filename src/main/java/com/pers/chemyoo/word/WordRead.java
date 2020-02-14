package com.pers.chemyoo.word;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.commons.io.FileUtils;
import org.apache.poi.POIXMLDocument;
import org.apache.poi.POIXMLTextExtractor;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

/**
 * 读取word文档生成建表语句
 * @author jianqing.liu
 * @since 2020年2月14日 下午12:08:05
 */
public class WordRead
{

	public static void main(String[] args)
	{
		WordRead tp = new WordRead();
		// .docx和doc文件的读取
		tp.readWord("C:\\Users\\chemyoo\\Desktop\\湖南电力\\数据库\\QR731801008湖南省电力质监信息化数据持久化设计说明书(BWX19191_V1.0.0__LJQ_2020.02.11).docx");
	}

	/**
	 * 读取word文件内容
	 * 
	 * @param path
	 * @return buffer
	 */
	public void readWord(String path)
	{
		String buffer = null;
		List<List<String>> cols = new ArrayList<>();
		try
		{
			if (path.endsWith(".doc"))
			{
				FileInputStream is = new FileInputStream(path);
				WordExtractor ex = new WordExtractor(is);
				buffer = ex.getText();
				is.close();
			}
			else if (path.endsWith("docx"))
			{
				OPCPackage opcPackage = POIXMLDocument.openPackage(path);
				POIXMLTextExtractor extractor = new XWPFWordExtractor(opcPackage);
				buffer = extractor.getText();
				opcPackage.close();
				InputStream is = new FileInputStream(path);  
				XWPFDocument doc = new XWPFDocument(is); 
				List<String> list = new ArrayList<>();  
		        List<XWPFParagraph>paras = doc.getParagraphs();  
		        for (XWPFParagraph graph : paras) {  
		            String text = graph.getParagraphText();  
		            String style = graph.getStyle();  
//		            if ("1".equals(style)) {  
//		              System.out.println(text+"--["+style+"]");  
//		            }else if ("2".equals(style)) {  
//		              System.out.println(text+"--["+style+"]");  
//		            }else 
		            if ("3".equals(style) && text.startsWith("HNDLZJ_")) {  
		              list.add(text);   
		            }else{  
		                continue;  
		            }  
		        }  
		        System.err.println(list);
		        File f = new File("D:", "text.txt");
		        FileUtils.writeStringToFile(f, buffer, "utf-8");
		        List<String> lines = FileUtils.readLines(f, "utf-8");
		        boolean start = false;
		        int index = -1;
		        List<String> col = null;
 		        for(String line : lines) {
		        	if(line.startsWith("名称	注释")) {
		        		if(index > -1) {
		        			cols.add(col);
		        		}
		        		start = true;
		        		index ++;
		        		col = new ArrayList<>();
		        		continue;
		        	}
		        	if("".equals(line.trim())) {
		        		start = false;
		        	}
		        	if(start) {
		        		System.err.println(line);
		        		col.add(line);
		        	}
		        }
 		        if(col != null && !col.isEmpty()) {
 		        	cols.add(col);
 		        }
 		        index = 0;
 		        File file = new File("D:", "auto create table.sql");
 		        if(file.exists()) {
 		        	file.delete();
 		        }
 		        for(String tab : list) {
 		        	StringBuilder sb = new StringBuilder();
 		        	String[] t = tab.split(" ");
 		        	sb.append("drop table if exists ").append(t[0]).append(";\r\n").append("\r\n");
 		        	sb.append("create table ").append(t[0]).append("(").append("\r\n");
 		        	int cc = 0;
 		        	for(String c : cols.get(index)) {
 		        		String[] ts = c.split("\t");
 		        		sb.append("\t").append(ts[0]).append(" ");
 		        		if(ts.length >= 4) {
 		        			sb.append(ts[2].replace("(n)", "(" + ts[3] + ")"));
 		        		} else {
 		        			sb.append(ts[2]);
 		        		}
 		        		if(cc == 0) {
 		        			sb.append(" not");
 		        		}
 		        		sb.append(" ").append("null comment '").append(ts[1]).append("',\r\n");
 		        		cc ++;
 		        	}
 		        	sb.append("constraint ").append(t[0]).append(" primary key clustered (ID)\r\n");
 		        	sb.append(") comment '").append(t[1]).append("';\r\n\r\n");
 		        	FileUtils.writeStringToFile(file, sb.toString(), true);
 		        	index ++;
 		        }
 		        f.delete();
			}
			else
			{
				System.out.println("此文件不是word文件！");
			}
		}
		catch (Exception e)
		{
			e.printStackTrace();
		}
	}

}
