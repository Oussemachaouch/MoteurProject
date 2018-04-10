package Moteur;


import Lib.ExcelDataConfig;

public class Moteur {

	public static void main(String[] args) {
		
	
			// Excel Configuration
				ExcelDataConfig excel = new ExcelDataConfig("C:\\Users\\ousse\\Desktop\\MoteurProject\\ExcelData.xlsx");
			
			 
				System.out.println("Iterations : ");
			
			
			// Make a map with ResultType from the Excel Sheet
				excel.addResultType();
			
			// Get the number of ResultTypes
				System.out.println(excel.getResultType().size());
			
			// Get all ResultTypes
				System.out.println(excel.getResultType());
			
			// Getting Data From EXCEL
				excel.GetDataByIteration(1);
			
				
			
		
	
		}
		
	}
	
		
	


