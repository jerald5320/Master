package com.hexa;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWrite {
	public static void main(String[] args) throws IOException {
		File FLoc = new File("C:\\Users\\jeral\\eclipse-workspace\\MavenClass\\Excel\\New.xlsx");
		Workbook w = new XSSFWorkbook();
		Sheet s = w.createSheet("Course");
		Row r= s.createRow(2);
		Cell c = r.createCell(1);
		c.setCellValue("java");
		
		FileOutputStream fOut =new FileOutputStream(FLoc);
		w.write(fOut);
		System.out.println("Done");
		
		
	}

}
