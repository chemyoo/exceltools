package com.pers.chemyoo.exceltools;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.logging.Logger;

import javax.swing.JFileChooser;
import javax.swing.filechooser.FileFilter;
import javax.swing.filechooser.FileSystemView;

import org.apache.commons.io.FileUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 * 将Excel文件中的sheet页签内容进行替换
 * @author chemyoo
 *
 */
public class ExcelUtils {
	
	public static final String OFFICE_EXCEL_2003_POSTFIX = ".xls";
	public static final String OFFICE_EXCEL_2010_POSTFIX = ".xlsx";
	public static Logger log = Logger.getGlobal();
	
	private static final String DO_REPLACE_FILE = "有关费用项目地区调整系数表.xlsx";
	
	private static boolean debug = true;
	
	// 存储替换Excel的内容
	private static final RWorkBook RW = new RWorkBook();
	
	private static class RWorkBook{
		
		private List<Map<Integer, Cell>> datas;
		
		private String selectPath = null;
		
		/** 读取替换文件的内容 */
		public List<Map<Integer, Cell>> getRWorkBookDatas(){
			if(datas == null) {
				InputStream is = ExcelUtils.class.getClassLoader().getResourceAsStream(DO_REPLACE_FILE);
				if(is == null) {
					is = ExcelUtils.class.getClassLoader().getResourceAsStream("/" + DO_REPLACE_FILE);
				}
				try {
					Workbook workbook = WorkbookFactory.create(is);
					datas = ExcelUtils.read(workbook, 0);
				} catch (IOException | InvalidFormatException e) {
					e.printStackTrace();
				}
			} 
			return datas;
		}
		
		public String getSelectPath() {
			return selectPath;
		}
		
		public void setSelectPath(File file) {
			this.selectPath = file.getAbsolutePath();
		}
		
	}
	
	/**
	 * 读取Excel转成WorkBook
	 * @param file
	 * @return
	 */
	public static Workbook readExcel(File file, final File errorPath) {
		if(file == null) {
			throw new RuntimeException("您没有选择文件...");
		}
		Workbook workbook = null;
		String path = file.getAbsolutePath();
		if (!isNotEmpty(path)) {
			return workbook;
		}
		try (InputStream is = new FileInputStream(path)){
			workbook = WorkbookFactory.create(is);
//				if (path.endsWith(OFFICE_EXCEL_2003_POSTFIX)) {
//					is = new FileInputStream(path);
//					workbook = new HSSFWorkbook(is);
//				} else if (path.endsWith(OFFICE_EXCEL_2010_POSTFIX)) {
//					is = new FileInputStream(path);
//					workbook = new XSSFWorkbook(is);
//				}
			
		} catch (Exception e) {
			log.info(String.format("读取Excel文件出错：【%s】\n", file.getAbsoluteFile()));
			e.printStackTrace();
			log.info(String.format("文件：【%s】处理失败，将拷贝到【%s】文件夹下", file.getName(), errorPath.getAbsoluteFile()));
			try {
				ExcelUtils.moveFile(file, errorPath);
//				FileUtils.copyFileToDirectory(file, new File(errorPath, ExcelUtils.getSubName(file)), true);
			} catch (IOException e1) {
				log.info("移动文件失败...");
			}
		}
		return workbook;
	}
	
	
	private static void moveFile(File orgin, File target) throws IOException {
		if(orgin.getAbsolutePath().contains(":/todo") || orgin.getAbsolutePath().contains(":\\todo")) {
			FileUtils.moveFileToDirectory(orgin, new File(target, ExcelUtils.getSubName(orgin)), true);
		} else {
			FileUtils.copyFileToDirectory(orgin, new File(target, ExcelUtils.getSubName(orgin)), true);
		}
	}
	
	private static boolean isNotEmpty(String str) {
		return !(str == null || str.trim().length() == 0);
	}
	
	public static File getFile() {
		JFileChooser fileChooser = new JFileChooser();//"F:/pic"
		if(!debug) {
			FileSystemView fsv = FileSystemView.getFileSystemView();  //注意了，这里重要的一句
			//设置最初路径为桌面路径              
			fileChooser.setCurrentDirectory(fsv.getHomeDirectory());
		} else {
			fileChooser.setCurrentDirectory(new File("H:/"));
		}
		fileChooser.setDialogTitle("请选择源文件夹");
		fileChooser.setApproveButtonText("确定");
		//只选择文件夹
		fileChooser.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES);
		//设置文件是否可多选
		fileChooser.setMultiSelectionEnabled(false);
		fileChooser.setAcceptAllFileFilterUsed(false);// 去掉显示所有文件的按钮
		fileChooser.setFileFilter(new FileFilter() {
			
			@Override
			public String getDescription() {
				return "请选择源文件夹";
			}
			
			@Override
			public boolean accept(File f) {
				String fileName = f.getName().toLowerCase();
				return f.isDirectory() || fileName.endsWith(".xls") || fileName.endsWith(".xlsx");
			}
		});
		if (fileChooser.showOpenDialog(null) == JFileChooser.APPROVE_OPTION) {
			return fileChooser.getSelectedFile();
		}
		return null;
	}
	
	public static File getPath(String title) {
		JFileChooser fileChooser = new JFileChooser();//"F:/pic"
		if(!debug) {
			FileSystemView fsv = FileSystemView.getFileSystemView();  //注意了，这里重要的一句
			//设置最初路径为桌面路径              
			fileChooser.setCurrentDirectory(fsv.getHomeDirectory());
		} else {
			fileChooser.setCurrentDirectory(new File("H:/"));
		}
		fileChooser.setDialogTitle(title);
		fileChooser.setApproveButtonText("确定");
		//只选择文件夹
		fileChooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
		//设置文件是否可多选
		fileChooser.setMultiSelectionEnabled(false);
		fileChooser.setAcceptAllFileFilterUsed(false);// 去掉显示所有文件的按钮
		fileChooser.setFileFilter(new FileFilter() {
			
			@Override
			public String getDescription() {
				return "选择文件";
			}
			
			@Override
			public boolean accept(File f) {
				String fileName = f.getName().toLowerCase();
				return f.isDirectory() || fileName.endsWith(".xls") || fileName.endsWith(".xlsx");
			}
		});
		if (fileChooser.showOpenDialog(null) == JFileChooser.APPROVE_OPTION) {
			return fileChooser.getSelectedFile();
		}
		return null;
	}
	
	private static String getSubName(File file) {
		String selectPath = RW.getSelectPath();
		String filePath = file.getParentFile().getAbsolutePath();
		return filePath.replace(selectPath, "");
	}
	
	public static void main(String[] args) throws IOException {
		// 读取源文件或源文件夹
		File file = ExcelUtils.getFile();
		if(file == null) {
			log.info("您没有选择文件，程序退出...");
			System.exit(0);
		}
		RW.setSelectPath(file);
		String rootPath = getRootPath(file);
		final File errorPath = new File(rootPath, "/error/"); 
		final File deskPath = new File(rootPath, "/success/"); 
		final File todo = new File(rootPath, "/todo/");
		if(!todo.exists()) todo.mkdir();
		ExcelUtils.handle(file, errorPath, deskPath);
		log.info(String.format("处理完成，处理成功文件已经保存到【%s】文件夹下", deskPath.getAbsoluteFile()));
		// 删除todo下的空文件夹
		ExcelUtils.doEnd(todo);
	}
	
	private static void doEnd(File todo) {
		File[] files = todo.listFiles();
		for(File f : files) {
			if(f.isDirectory()) {
				ExcelUtils.doEnd(f);
			}
			f.delete();
		}
	}
	
	private static String getRootPath(File file) {
		return file.getAbsolutePath().substring(0, 1) + ":";
	}
	
	/**
	 * 递归调用
	 * @param file
	 * @param repalceBook
	 */
	private static void handle(File file, final File errorPath, final File deskPath) {
		if(file.isDirectory()) {
			File[] files = file.listFiles();
			for(File f : files) {
				try {
					ExcelUtils.handle(f, errorPath, deskPath);
				} catch (RuntimeException e) {
					log.info(e.getMessage());
					log.info(String.format("读取Excel文件出错：【%s】\n", f.getAbsoluteFile()));
					log.info(String.format("文件：【%s】处理失败，将拷贝到【%s】文件夹下", f.getName(), errorPath.getAbsoluteFile()));
					try {
						ExcelUtils.moveFile(f, errorPath);
//						FileUtils.copyFileToDirectory(f, new File(errorPath, ExcelUtils.getSubName(file)), true);
					} catch (IOException e1) {
						log.info("移动文件失败...");
					}
				}
			}
		} else if(file.getName().endsWith(".xls") || file.getName().endsWith(".xlsx")) {
			Workbook book = ExcelUtils.readExcel(file, errorPath);
			if(book == null) return;
			int sheetIndexs = book.getNumberOfSheets();
			boolean hasItem = false;
			for(int i = 0; i < sheetIndexs; i ++) {
				String sheetName = book.getSheetName(i);
				if(sheetName.contains("有关费用项目地区调整系数")) {
					hasItem = !hasItem;
					Workbook newbook = ExcelUtils.repalce(book, i, RW.getRWorkBookDatas());
					File newFile = new File(deskPath + ExcelUtils.getSubName(file), file.getName());
					if(!newFile.exists()) {
						newFile.getParentFile().mkdirs();
					}
					try(OutputStream out = new FileOutputStream(newFile)){
						//刷新公式
//						newbook.setForceFormulaRecalculation(true);
//	                    HSSFFormulaEvaluator.evaluateAllFormulaCells(newbook);
						newbook.write(out);
					} catch (FileNotFoundException e) {
						e.printStackTrace();
					} catch (IOException e) {
						e.printStackTrace();
					}
					if(file.getAbsolutePath().contains(":/todo") || file.getAbsolutePath().contains(":\\todo")) {
						FileUtils.deleteQuietly(file);
					}
					
				}
			}
			if(!hasItem) {
				try {
					FileUtils.copyFileToDirectory(file, new File(deskPath + ExcelUtils.getSubName(file)), true);
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		} 
//		else {
//			try {
//				FileUtils.copyFileToDirectory(file, new File(deskPath + ExcelUtils.getSubName(file)), true);
//			} catch (IOException e) {
//				e.printStackTrace();
//			}
//		}
	}
	
	// 读取替换表格的内容
	private static List<Map<Integer,Cell>> read(Workbook book, int indexOfSheet) {
		int sheetIndexs = book.getNumberOfSheets();
		if(indexOfSheet > sheetIndexs) {
			throw new RuntimeException("indexOfSheet out of range [0," + sheetIndexs + ")");
		}
		Sheet sheet = book.getSheetAt(indexOfSheet);
		int start = sheet.getFirstRowNum();
		int end = sheet.getLastRowNum();
		List<Map<Integer,Cell>> datas = new ArrayList<>();
		for(;start < end; start++) {
			Row row = sheet.getRow(start);
			int rows = row.getLastCellNum();
			Map<Integer,Cell> map = new HashMap<>();
			for(int j = 0; j < rows; j ++) {
				Cell obj = row.getCell(j);
				if(obj != null)
					map.put(j, obj);
				else {
					map.put(j, row.createCell(j));
				}
			}
			datas.add(map);
		}
		return datas;
	}
	
	/**
	 * 进行替换值
	 * @param book
	 * @param indexOfSheet
	 * @param repalceDatas
	 * @return
	 */
	private static Workbook repalce(Workbook book, int indexOfSheet, List<Map<Integer,Cell>> repalceDatas) {
		int sheetIndexs = book.getNumberOfSheets();
		if(indexOfSheet > sheetIndexs) {
			throw new RuntimeException("indexOfSheet out of range [0," + sheetIndexs + ")");
		}
		Sheet sheet = book.getSheetAt(indexOfSheet);
		int rowIndex = 0;
		try {
			for(Map<Integer,Cell> map : repalceDatas) {
				Row row = sheet.getRow(rowIndex);
				if(row == null) {
					row = sheet.createRow(rowIndex);
				}
				for(Map.Entry<Integer,Cell> entry : map.entrySet()) {
					Cell cell = row.getCell(entry.getKey());
					Cell rCell = entry.getValue();
					if(cell == null) {
						cell = row.createCell(entry.getKey());
						CellStyle style = book.createCellStyle();
						style.cloneStyleFrom(rCell.getCellStyle());
						cell.setCellStyle(style);
					}
					if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
						cell.setCellValue(rCell.getNumericCellValue());
					} else if (cell.getCellType() == Cell.CELL_TYPE_BOOLEAN) {
						cell.setCellValue(rCell.getBooleanCellValue());
					} else if (cell.getCellType() == Cell.CELL_TYPE_FORMULA) {
						cell.setCellValue(rCell.getCellFormula());
					} else if(cell.getCellType() == Cell.CELL_TYPE_ERROR) {
						cell.setCellErrorValue(rCell.getErrorCellValue());
					} else if(rCell.getCellType() != Cell.CELL_TYPE_STRING){
						cell.setCellType(rCell.getCellType());
						if (rCell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
							cell.setCellValue(rCell.getNumericCellValue());
						} else if (rCell.getCellType() == Cell.CELL_TYPE_BOOLEAN) {
							cell.setCellValue(rCell.getBooleanCellValue());
						} else if (rCell.getCellType() == Cell.CELL_TYPE_FORMULA) {
							cell.setCellValue(rCell.getCellFormula());
						} else if(rCell.getCellType() == Cell.CELL_TYPE_ERROR) {
							cell.setCellErrorValue(rCell.getErrorCellValue());
						} else {
							cell.setCellValue(rCell.getStringCellValue());
						}
					} else {
						cell.setCellValue(rCell.getStringCellValue());
					}
					 // 使用evaluateFormulaCell对函数单元格进行强行更新计算
//					book.getCreationHelper().createFormulaEvaluator().evaluateFormulaCell(cell);
				}
				rowIndex ++;
			}
		} catch (Exception e) {
			throw new RuntimeException("页签【" + sheet.getSheetName() + "】第" + (rowIndex + 1) + "行出现错误：" + e.getMessage());
		}
		return book;
	}
}
