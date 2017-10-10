package main;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Reader {

	private String filePath;
	private String fileName;
	private String sheetName;
	private Sheet sheet;
	private Workbook workBook;
	
	public Reader() {
		super();
	}
	public Reader(String filePath, String fileName, String sheetName) {
		super();
		this.filePath = filePath;
		this.fileName = fileName;
		this.sheetName = sheetName;
	}
	public String getFilePath() {
		return filePath;
	}
	public void setFilePath(String filePath) {
		this.filePath = filePath;
	}
	public String getFileName() {
		return fileName;
	}
	public void setFileName(String fileName) {
		this.fileName = fileName;
	}
	public String getSheetName() {
		return sheetName;
	}
	public void setSheetName(String sheetName) {
		this.sheetName = sheetName;
	}
	public Sheet getSheet() {
		return sheet;
	}
	public Workbook getWorkBook() {
		return workBook;
	}	
	
	public void Read() throws IOException {
		File file = new File(filePath + "\\" + fileName);
		FileInputStream fis = new FileInputStream(file);
		
		Workbook wb = null;
		
		String fex = fileName.substring(fileName.indexOf("."));
		
		if(fex.equals(".xls")) {
			wb = new HSSFWorkbook(fis);
		}else if(fex.equals(".xlsx")) {
			wb = new XSSFWorkbook(fis);
		}
		
		this.workBook = wb;
		this.sheet = wb.getSheet(sheetName);
		System.out.println(this.sheet);
	}
	
	public void close() throws IOException {
		this.workBook.close();
	}
	
}
