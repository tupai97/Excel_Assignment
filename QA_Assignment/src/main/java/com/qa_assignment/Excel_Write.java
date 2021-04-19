package com.qa_assignment;

import java.io.FileOutputStream;
import java.io.IOException;
import java.text.ParseException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel_Write {
	Excel_Read excelRead = new Excel_Read();

	public final void Excel_data_write() throws IOException, ParseException {

		String exl_File_final = "./ExcelData/Final_excelfile.xlsx";  
		XSSFWorkbook workbook = new XSSFWorkbook();   
		XSSFSheet sheet = workbook.createSheet("Excel_Data_final");
		String[][] dummy_val  = excelRead.excel_data_extraction().clone();

		int dataRows = dummy_val.length;
		int dataColumn = dummy_val[0].length;

		for(int i=0 ; i<dataRows ;i++) {
			Row row = sheet.createRow(i);
			for(int j=0 ; j<dataColumn ; j++) {
				String fillData = dummy_val[i][j];
				Cell cell = row.createCell(j);
				cell.setCellValue(fillData);	
			}
		}
		FileOutputStream fileout = new FileOutputStream(exl_File_final);
		workbook.write(fileout);
		fileout.close();
	}	
}
