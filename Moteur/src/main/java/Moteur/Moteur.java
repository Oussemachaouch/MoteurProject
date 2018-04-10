package Moteur;

import java.io.File;
import java.io.FileInputStream;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import lib.ExcelDataConfig;

public class Moteur {

	public static void main(String[] args) {
		
		
		// TODO Auto-generated method stub
		XSSFWorkbook wb;
		XSSFSheet sheet1;

		
		try {
			
			File src = new File("C:\\Users\\ousse\\Desktop\\MoteurProject\\ExcelData.xlsx");
			FileInputStream fis=new FileInputStream(src);
			wb = new XSSFWorkbook(fis);
			
			ExcelDataConfig excel = new ExcelDataConfig("C:\\Users\\ousse\\Desktop\\Excel\\ExcelData.xlsx");
			
			System.out.println("Iterations : ");
			
			
			// Make a map with ResultType from the Excel Sheet
				excel.addResultT();
			
			// Get the number of ResultTypes
				System.out.println(excel.getResultT().size());
			
			// Get all ResultTypes
				System.out.println(excel.getResultT());
			
			// Getting Data From EXCEL
				excel.GetDataByIteration(1);
			
				
			} catch (Exception e) {
			// TODO Auto-generated catch block
			e.getMessage();
		}
		
	
		}
		
	}
	
		
	


