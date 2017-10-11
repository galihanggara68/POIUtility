package main;

import java.io.IOException;
import java.io.PrintStream;

import com.galih.poi.Reader;


public class Main {

	public static void main(String[] args) throws IOException{
		PrintStream p = new PrintStream(System.out);
		Reader r = new Reader("D:/Selenium/", "test.xlsx", "Test");
		r.setRegion(Reader.DEFAULT_REGION);
		r.Read();
		for(int i = 0; i < r.getResults().length; i++){
			for(int j = 0; j < r.getResults()[i].length; j++){
				p.print(r.getResults()[i][j].getStringCellValue() + "|");
			}
			p.print("\n");
		}
	}
	
}
