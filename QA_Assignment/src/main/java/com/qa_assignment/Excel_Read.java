package com.qa_assignment;

import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.time.temporal.ChronoUnit;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.HashSet;
import java.util.Iterator;
import java.util.Set;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
public class Excel_Read  {
	@SuppressWarnings({ })
	public String[][] excel_data_extraction() throws IOException, ParseException {
		//File file = new File("C:\\Users\\91858\\Desktop\\computer\\sample_excelfile.xlsx");         //creating a new file instance  
		//FileInputStream fis = new FileInputStream(file);                                            //obtaining bytes from the file  
		String exl_File = "./ExcelData/sample_excelfile.xlsx";                                        // Adding file in the same project for Efficiency 
		XSSFWorkbook wb = new XSSFWorkbook(exl_File);                                                 //creating Workbook instance that refers to .xlsx file 
		XSSFSheet sheet = wb.getSheetAt(0);                                                           //creating a Sheet object to retrieve particular sheet from workbook  	
		
		Iterator<Row> iterator = sheet.iterator();     	    
		SimpleDateFormat sdf = new SimpleDateFormat("HH:mm:ss");
		
		Calendar c = Calendar.getInstance();
		DateTimeFormatter myformattype1 = DateTimeFormatter.ofPattern("d-MMM-yyyy");
		ArrayList<ArrayList<String> > dummyArray = new ArrayList<ArrayList<String> >();
		ArrayList<String> values = null;	
		
		int rowCount= sheet.getPhysicalNumberOfRows();
		Set<String> vci_code_set= new HashSet<String>();

		//Adding new column name at row0 and column12  by name - "Time taken(days)"
		while (iterator.hasNext()) {
			Row currentRow = iterator.next();
			Cell cell = currentRow.createCell(currentRow.getLastCellNum(), Cell.CELL_TYPE_STRING);
			if(currentRow.getRowNum() == 0) {
				cell.setCellValue("Time taken(days)");
			}
			else {
				cell.setCellType(Cell.CELL_TYPE_BLANK);
			}
		}
		for (int row=0; row<rowCount; row++) {
			values  = new ArrayList<String>();
			if (vci_code_set.contains(sheet.getRow(row).getCell(0).toString())==false)
			{
				//Adding VCI_Code to the list to check for duplicates
				vci_code_set.add(sheet.getRow(row).getCell(0).toString());
				int colCount= sheet.getRow(row).getLastCellNum();
				for (int col=0; col<colCount; col++) {
					Cell cell= sheet.getRow(row).getCell(col);
					if(cell.getCellType() == Cell.CELL_TYPE_STRING ){
						values.add(sheet.getRow(row).getCell(col).toString());
					}
					else {
						if (cell.getCellType()==Cell.CELL_TYPE_NUMERIC  && DateUtil.isCellDateFormatted(cell)) {
							Date date= cell.getDateCellValue();	
							SimpleDateFormat formatTime= new SimpleDateFormat("HH:mm:ss");
							SimpleDateFormat formatYearOnly= new SimpleDateFormat("yyyy");
							String dateStamp= formatYearOnly.format(date);
							if (dateStamp.equals("1899")) {
								if(formatTime.format(date).toString()!=null) {
									values.add(formatTime.format(date).toString());
								}			
							} 
							else {
								String timeStamp =formatTime.format(date);
								if (timeStamp.equals("00:00:00")) {
									if(cell.toString()!=null) {
										values.add(cell.toString());
									}	
								}
							}
						}
					}
				}
			}
			else {
				continue;
			}
			dummyArray.add(values);
			System.out.println(" ");
		}

		//Update the column name from "Time Ingested (MNL Time)" to "AEST" & change the value of MNL time to AEST accordingly
		for(int i = 0; i < dummyArray.size(); i++){
			if (dummyArray.get(i).get(9).trim().equals("Time Ingested (MNL Time)") ) {
				dummyArray.get(i).set(9, "AEST");	
			}
			else {
				Date date = sdf.parse(dummyArray.get(i).get(9));
				c.setTime(date);
				c.add(Calendar.HOUR, +4);
				c.add(Calendar.MINUTE, 20);
				c.add(Calendar.SECOND, 00);
				Date currentDatePlusOne = c.getTime();	    
				sdf.format(currentDatePlusOne);
				dummyArray.get(i).set(9, sdf.format(currentDatePlusOne).toString());
			}
		}

		//Create a temporary array for storage of "Number Of Days Between - (Date received - date of decision)"
		String[] numberOfDaysBtw = new String[dummyArray.size()];

		// minus column 1 from column 7 & store in array numberOfDaysBtw
		for (int i = 1 ; i<dummyArray.size() ;i++ ) {
			LocalDate dateReceived = LocalDate.parse(dummyArray.get(i).get(1),myformattype1);
			LocalDate dateOfDecision = LocalDate.parse(dummyArray.get(i).get(7),myformattype1);
			long noOfDaysBetween =  ChronoUnit.DAYS.between(dateOfDecision , dateReceived);
			numberOfDaysBtw[i] = Long.toString(noOfDaysBetween);
		}
		
		//putting the values of dummyArray into final_array 
		String[][] final_array = new String[dummyArray.size()][dummyArray.get(0).size()];
		for(int i = 0; i < dummyArray.size(); i++){
			for(int j = 0; j < dummyArray.get(i).size(); j++){			
				final_array[i][j]=dummyArray.get(i).get(j);
			}
		}
		//Adding value of numberOfDaysBtw array into column12 of final_array
		for(int i = 1; i < dummyArray.size(); i++){
			final_array[i][12] = numberOfDaysBtw[i];
			final_array[i][11] = " ";
			System.out.println();
		}		
		// Printing the final array for the output 
//				for(int i = 0; i < dummyArray.size(); i++){
//					for(int j = 0; j < dummyArray.get(0).size(); j++){			
//						System.out.print(final_array[i][j]+"\t\t\t");
//					}
//					System.out.println();
//				}	
				return final_array;				
	}
}








