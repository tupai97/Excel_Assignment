package com.qa_assignment;

public class Main_Tester {

		public static void main(String[] args) {	
			Excel_Write Ew = new Excel_Write();
			try 
			{
				Ew.Excel_data_write();
			}
			catch (Exception e){
				e.printStackTrace();
			}
		}
}
