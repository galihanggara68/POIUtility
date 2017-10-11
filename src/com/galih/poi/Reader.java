package com.galih.poi;
/*
 * Author : Muchamad Galih Anggara
 * Name : POI Reader Utility
 * Description : Reads excel file and gets all results from specific sheet
 * */
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Reader {

	private String filePath;
	private String fileName;
	private String sheetName;
	private Sheet sheet;
	private Workbook workBook;
	private int[] region = new int[] {0, 0, 0, 0};
	public static final int[] DEFAULT_REGION = new int[] {0, 0, 0, 0};
	private Cell[][] results;
	
	public Reader() {
		
	}
	public Reader(String filePath, String fileName, String sheetName) {
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
	public void setRegion(int[] region){
		this.region = region;
	}
	public void setRegion(int rowStart, int rowEnd, int cellStart, int cellEnd){
		this.region = new int[] {rowStart, rowEnd, cellStart, cellEnd};
	}
	public int[] getRegion(){
		return this.region;
	}
	public Sheet getSheet() {
		return sheet;
	}
	public Workbook getWorkBook() {
		return workBook;
	}	
	public Cell[][] getResults(){
		return this.results;
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
		Sheet sheet = wb.getSheet(sheetName);
		int rowCount = sheet.getLastRowNum() - sheet.getFirstRowNum();
		if(region == DEFAULT_REGION ){
			region[1] = rowCount;
			region[3] = sheet.getRow(0).getLastCellNum()-1;
		}
		int inc = 0;
		results = new Cell[(region[1]-region[0])+1][(region[3]-region[2])+1];
		for(int i = region[0]; i <= region[1]; i++){
			Row row = sheet.getRow(i);
			for(int j = region[2]; j <= region[3]; j++){
				results[inc][j] = row.getCell(j);
			}
			inc++;
		}
		this.workBook = wb;
		this.sheet = wb.getSheet(sheetName);
	}
	
	public void close() throws IOException {
		this.workBook.close();
	}
	
}
