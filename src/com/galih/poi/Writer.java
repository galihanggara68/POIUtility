package com.galih.poi;
/*
 * Author : Muchamad Galih Anggara
 * Name : POI Writer Utility
 * Description : Writes or replaces excel file to specific sheet
 * */
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Writer {
	private String filePath;
	private String fileName;
	private String sheetName;
	private int[] region = new int[] {0, 0, 0};
	
	//Constants
	public static final int[] DEFAULT_REGION = new int[] {0, 0, 0};
	
	public Writer() {
	}
	
	public Writer(String filePath, String fileName, String sheetName) {
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
	
	public void setRegion(int[] region) {
		this.region = region;
	}
	
	public void setRegion(int rowStart, int cellStart, int cellEnd) {
		this.region = new int[] {rowStart, cellStart, cellEnd};
	}
	
	public int[] getRegion(){
		return this.region;
	}
	
	public void Write(String[] datas) throws IOException {
		File file = new File(filePath + "\\" + fileName);
		FileInputStream fis = new FileInputStream(file);
		
		Workbook wb = null;
		
		String fex = fileName.substring(fileName.indexOf("."));
		
		if(fex.equals(".xls")) {
			wb = new HSSFWorkbook(fis);
		}
		if(fex.equals(".xlsx")) {
			wb = new XSSFWorkbook(fis);
		}
		
		Sheet sheet = wb.getSheet(sheetName);
		int rowCount = sheet.getLastRowNum() - sheet.getFirstRowNum();
		if(region == DEFAULT_REGION){
			region[2] = sheet.getRow(0).getLastCellNum();
		}
		Row nr = sheet.createRow(rowCount + 1);
		for(int i = region[1]; i <= region[2]; i++) {
				Cell cell = nr.createCell(i);
				cell.setCellValue(datas[i-region[1]]);
		}
		fis.close();
		FileOutputStream fos = new FileOutputStream(file);
		wb.write(fos);
		wb.close();
		fos.close();
	}
	
	public void Replace(String[] datas) throws IOException {
		File file = new File(filePath + "\\" + fileName);
		FileInputStream fis = new FileInputStream(file);
		
		Workbook wb = null;
		
		String fex = fileName.substring(fileName.indexOf("."));
		
		if(fex.equals(".xls")) {
			wb = new HSSFWorkbook(fis);
		}
		if(fex.equals(".xlsx")) {
			wb = new XSSFWorkbook(fis);
		}
		
		Sheet sheet = wb.getSheet(sheetName);
		int rowCount = sheet.getLastRowNum() - sheet.getFirstRowNum();
		if(region == DEFAULT_REGION){
			region[2] = sheet.getRow(0).getLastCellNum();
		}
		Row nr = sheet.getRow(region[0]);
		for(int i = region[1]; i <= region[2]; i++) {
				if(nr.getCell(i) == null){
					nr.createCell(i).setCellValue(datas[i - region[1]]);
				}else{
					nr.getCell(i).setCellValue(datas[i - region[1]]);
				}
		}
		fis.close();
		FileOutputStream fos = new FileOutputStream(file);
		wb.write(fos);
		wb.close();
		fos.close();
	}
}
