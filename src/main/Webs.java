package main;

import java.io.IOException;
import java.util.Scanner;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

/*
 * Input file and output file must be a different file, same file will cause data corrupt.
 * In this case, i use 'test.xlsx' file for Input Stream, and 'out.xlsx' for Input Stream.
 * */

public class Webs {

	public static void main(String[] args) throws InterruptedException, IOException{
		System.setProperty("webdriver.chrome.driver", "C:/Users/Geeksfarm/Desktop/selenium-java-3.6.0/chromedriver.exe");
		WebDriver driver = new ChromeDriver();
		driver.get("http://www.store.demoqa.com");
		driver.manage().window().maximize();
		
		//Create 
		String filePath = System.getProperty("user.dir") + "/src/main";
		String fileName;
		String sheetName;
		Scanner sc = new Scanner(System.in);
		System.out.println("Masukan Nama File : ");
		fileName = sc.nextLine();
		System.out.println("Masukan Nama Sheet : ");
		sheetName = sc.nextLine();
		Writer wx = new Writer(filePath, fileName, sheetName);
		Reader rx = new Reader(filePath, fileName, sheetName);
		rx.Read();
		Sheet fsh = rx.getSheet();
		int rowNum = fsh.getLastRowNum()-fsh.getFirstRowNum();
		
		//Looping for input data to form and write the report
		for(int i = 1; i <= rowNum; i++){
			
			//Get current Row
			Row rw = fsh.getRow(i);
				driver.findElement(By.className("account_icon")).click();
				driver.findElement(By.id("log")).sendKeys(rw.getCell(0).getStringCellValue());
				driver.findElement(By.id("pwd")).sendKeys(rw.getCell(1).getStringCellValue());
				driver.findElement(By.id("login")).click();
				Thread.sleep(5000);
				if(driver.findElement(By.className("response")).getText().substring(0 ,driver.findElement(By.className("response")).getText().indexOf(":")).equalsIgnoreCase("error")){
					System.out.println("Erro Bro !");
					//Write status for every single user
					wx.setRegion(i, 2, 3);
					wx.Replace(new String[] {"Failed", driver.findElement(By.className("response")).getText()});
				}else{
					System.out.println("Berhasil Bro !");
					//Write status for every single user
					wx.setRegion(i, 2, 3);
					wx.Replace(new String[] {"Pass", driver.findElement(By.className("response")).getText()});
				}
		}
	}
	
}
