package Lib;


import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import ExcelData.ExcelResultType;


public class ExcelDataConfig {
	
	//Variable Globale
	
	XSSFWorkbook wb;
	XSSFSheet sheet1;
	ExcelResultType ResultType;
	private Map<Integer,String> ResultT =new HashMap<Integer,String>();  
	
	
		public ExcelDataConfig(String excelPath){
		
			try {
				File src = new File(excelPath);
				FileInputStream fis=new FileInputStream(src);
				wb = new XSSFWorkbook(fis);
				
				
				if(src.isFile() && src.exists()) {
				     System.out.println("File open successfully.");
				  } else {
				     System.out.println("Error to open  file.");
				  }
			}
			 catch (Exception e) {
				// TODO Auto-generated catch block
				e.getMessage();
			}
			}

		public Map<Integer, String> getResultT() {
			return ResultT;
		}

		public void setResultT(Map<Integer, String> resultT) {
			ResultT = resultT;
		}
		
		
		public Map<Integer, String> addResultT() {
			int j=0;
			
			while (j<ResultType.values().length)
			{
				
				String val = ResultType.values()[j].getDesc();
				j++;
				ResultT.put(j, val);
			}
		
			return ResultT;
		}


		public int getIteratorRowPosition(int sheetNumber){
			sheet1= wb.getSheetAt(sheetNumber);
			XSSFRow row; 
			XSSFCell cell;
			Iterator rows = sheet1.rowIterator();

			while (rows.hasNext())
			{
				row=(XSSFRow) rows.next();
				Iterator cells = row.cellIterator();
				
				while (cells.hasNext())
				{ 
					cell=(XSSFCell) cells.next();
			
					if (cell.getCellType() == XSSFCell.CELL_TYPE_STRING && cell.getStringCellValue().equals("ITERATOR"))
					{
						return cell.getRowIndex()+1;
					}
			
				}
				
			}
		
			return 0;
		}
		
		
		public int getIteratorColumnPosition(int sheetNumber){
			sheet1= wb.getSheetAt(sheetNumber);
			
			XSSFRow row; 
			XSSFCell cell;
		
			Iterator rows = sheet1.rowIterator();

			while (rows.hasNext())
			{
				row=(XSSFRow) rows.next();
				Iterator cells = row.cellIterator();
				
				while (cells.hasNext())
				{ 
					cell=(XSSFCell) cells.next();
			
					if (cell.getCellType() == XSSFCell.CELL_TYPE_STRING && cell.getStringCellValue().equals("ITERATOR"))
					{
						return cell.getColumnIndex();
					}
	
				}
				
			}
		
			return 0;
		}
		
		
		public double getIterator(int sheetNumber){
			sheet1= wb.getSheetAt(sheetNumber);
			XSSFRow row; 
			XSSFCell cell;
		
			Iterator rows = sheet1.rowIterator();
			
			while (rows.hasNext())
			{
				row=(XSSFRow) rows.next();
				Iterator cells = row.cellIterator();
				
				while (cells.hasNext())
				{ 
					cell=(XSSFCell) cells.next();
			
				
					if (cell.getCellType() == XSSFCell.CELL_TYPE_STRING && cell.getStringCellValue().equals("ITERATOR"))
					{
						
						System.out.print(cell.getStringCellValue()+" ");
						System.out.println(getNumericData(1, cell.getRowIndex()+1, cell.getColumnIndex()));
						
						return getNumericData(1, cell.getRowIndex()+1, cell.getColumnIndex());
					}
					
					
				}
				
			}
		return 0.0;
		}

		
		public void GetDataByIteration(int sheetNumber){
			
				sheet1= wb.getSheetAt(sheetNumber);
		//Get Iterator
				double k = getIterator(sheetNumber);
						
				sheet1= wb.getSheetAt(1);
				XSSFRow row=sheet1.createRow(getIteratorRowPosition(sheetNumber));
				XSSFCell cell=row.createCell(getIteratorColumnPosition(sheetNumber));
		
		// Loop Iteration
				for  ( double j =  k ; j<= 12; j++)
					{
						System.out.println("-----------");
						System.out.println("Iteration Number "+j);
						System.out.println("-----------");			
		// Update CellValue	
					cell.setCellValue(j);
						
		// Update Sheet				
					wb.getCreationHelper().createFormulaEvaluator().evaluateAll();
						
		// Get Data From Excel				
					GetAllData(1);
								
					}
		}
		
		public String GettingType(int i){
			String var =ResultT.get(i);
			return var;
		}
	
		public void GetAllData(int sheetNumber){
			sheet1= wb.getSheetAt(sheetNumber);
			
			XSSFRow row; 
			XSSFCell cell;
			Iterator rows = sheet1.rowIterator();
		    
			while (rows.hasNext())
			{
				row=(XSSFRow) rows.next();
				Iterator cells = row.cellIterator();
				
				while (cells.hasNext())
				{ 
					cell=(XSSFCell) cells.next();
					for (int i=0 ; i<=ResultT.size() ; i++)
					{		
						if (cell.getCellType() == XSSFCell.CELL_TYPE_STRING && cell.getStringCellValue().equals(ResultT.get(i)))
							{
								System.out.print(cell.getStringCellValue()+" ");
								System.out.println(getNumericData(1, cell.getRowIndex(), cell.getColumnIndex()+1));
							}
					}
					
				}
				
			 }
		
		}
		
		public int getRowNumber(int sheetNumber) {
			sheet1= wb.getSheetAt(sheetNumber);
			int rowCount = sheet1.getLastRowNum()-sheet1.getFirstRowNum()+1;
			return rowCount;
        }
		
		
		public double getNumericData(int sheetNumber,int row ,int column){
			sheet1= wb.getSheetAt(sheetNumber);
			double data =sheet1.getRow(row).getCell(column).getNumericCellValue();
			return data;
		}
		
		
		
		public String getType(int sheetNumber,int row ,int column){
			sheet1= wb.getSheetAt(sheetNumber);
			String data =sheet1.getRow(row).getCell(column).getStringCellValue();
			return data;
		}
		
		public Map<Integer,String> getMapData(int sheetNumber,int column){
			
			
			sheet1= wb.getSheetAt(sheetNumber);
			
			for(int i = 0; i <= getRowNumber(0)-1; i++)
			{
				
				String data =sheet1.getRow(i).getCell(column).getStringCellValue();
				ResultT.put(i, data);
			
			}
			
		return ResultT;
		}
		
		
		/*public int getIterationNumber(int sheetNumber,int row ,int column)
		{
			sheet1= wb.getSheetAt(sheetNumber);
			int data =sheet1.getRow(row).getCell(column).getCachedFormulaResultType();
			return data;
		}
		
		public boolean getCellV(int sheetNumber,int row ,int column)
		{
			sheet1= wb.getSheetAt(sheetNumber);
			sheet1.getRow(row).getCell(column).getCellFormula();
			try {
				FileOutputStream output_file =new FileOutputStream(new File("C:\\Users\\ousse\\Desktop\\Excel\\ExcelData.xlsx"));  //Open FileOutputStream to write updates
				
				wb.write(output_file);
				output_file.close();
			} catch (Exception e) {
				// TODO Auto-generated catch block
				e.getMessage();
			}
			
		}
		return true;*/
}
