package com.Excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;


public class ReadExcel {
//@Test
public void read() throws IOException {
	File fl=new File("D:\\Read_Excel\\Read.xlsx");
	FileInputStream fls=new FileInputStream(fl);
	XSSFWorkbook wb=new XSSFWorkbook(fls);
	XSSFSheet sheet=wb.getSheetAt(0);
	System.out.println(sheet.getRow(1).getCell(1).getStringCellValue());
	int row=sheet.getLastRowNum();
	DataFormatter format=new DataFormatter();
	for(int i=0;i<row;i++) {
		Row r=sheet.getRow(i);
		for(int j=0;j<r.getLastCellNum();j++) {
			System.out.println(format.formatCellValue(r.getCell(j)));
		}
		
	}
	
}
}
