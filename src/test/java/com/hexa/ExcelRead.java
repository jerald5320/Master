package com.hexa;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelRead {
public static void main(String[] args) throws IOException {
	File excelLoc =new File("C:\\Users\\jeral\\eclipse-workspace\\MavenClass\\Excel\\Book2.xlsx");
	FileInputStream FIn= new FileInputStream(excelLoc);
	Workbook w = new XSSFWorkbook(FIn);
	Sheet s= w.getSheet("Data");
	
	for (int i = 0; i < s.getPhysicalNumberOfRows(); i++) {
		Row r = s.getRow(i);
		for (int j = 0; j < r.getPhysicalNumberOfCells(); j++) {
			Cell c = r.getCell(j);
		
			int type = c.getCellType();
			
			if (type==1) {
				String sValue = c.getStringCellValue();
				System.out.print(sValue+"\t");
			
				
				
				
				
				
			}
			else if (type==0) {
				if (DateUtil.isCellDateFormatted(c)) {
					Date date = c.getDateCellValue();
					SimpleDateFormat self=new SimpleDateFormat("dd/mm/yy");
					String dValue = self.format(date);
					System.out.print(dValue+"\t");
							
				}
				else {
					double d = c.getNumericCellValue();
					long l=(long)d;
					String nValue = String.valueOf(l);
					System.out.print(nValue+"\t");
					
				}
			}
		}
		System.out.println();
	}
	
	
	
	
	
	
		
	
	
}
}
