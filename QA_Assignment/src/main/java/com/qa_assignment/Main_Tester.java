package com.qa_assignment;

import java.util.logging.Level;

import org.apache.log4j.ConsoleAppender;
import org.apache.log4j.Logger;
import org.apache.log4j.PatternLayout;

public class Main_Tester {
 
	 static final Logger logger = Logger.getLogger(Main_Tester.class);
	
	  
	 public static void main(String[] args) {	
		 	logger.info("Entering the Execute method ");
			Excel_Write Ew = new Excel_Write();
			try 
			{
				logger.info("save successful" + Ew.toString());
				Ew.Excel_data_write();
			}
			catch (Exception e){
				logger.error("Error message are"+ e.getMessage());
				e.printStackTrace();
			}
			
			
			
			
		}
}
