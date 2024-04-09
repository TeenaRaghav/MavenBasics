package com.training.day9;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadXl {
	static XSSFWorkbook excelWorkbook;
	static XSSFSheet excelSheet;

	public static void main(String[] args) {
		try {
// define the path
//		String userDir = System.getProperty("user.dir");
//		String fileseperator = System.getProperty("file.separator");
//		
			File filepath = new File("C:\\feb2024\\java\\ReadExcel.xlsx");

//		add this file into fileinput stream
			FileInputStream excelfile = new FileInputStream(filepath);

//		file inputstream should be converted into workbook
			excelWorkbook = new XSSFWorkbook(excelfile);
			excelSheet = excelWorkbook.getSheet("Sheet2");
//			System.out.println(excelSheet.getRow(2).getCell(1));

			int totalRows = excelSheet.getLastRowNum();
			for (int i = 0; i <= totalRows; i++) {

				for (int j = 0; j < 2; j++) {
					System.out.print(excelSheet.getRow(i).getCell(j));
					System.out.print("\t");
				}
				System.out.println();
			}
			excelfile.close();
		} catch (Exception e) {
			System.out.println(e);
		}

	}
}
