package Lib;


import java.io.File;
import java.io.FileInputStream;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import ExcelData.ExcelResultType;


public class ExcelDataConfig {
	
	
	XSSFWorkbook wb;
	XSSFSheet sheet1;
	ExcelResultType ResultType;
	private Map<Integer,String> ResultTypeMap =new HashMap<Integer,String>();  
	
	
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

		public Map<Integer, String> getResultType() {
			return ResultTypeMap;
		}

		public void setResultType(Map<Integer, String> ResultTypeMap) {
			ResultTypeMap = ResultTypeMap;
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
		
		public double getNumericData(int sheetNumber,int row ,int column){
			sheet1= wb.getSheetAt(sheetNumber);
			double data =sheet1.getRow(row).getCell(column).getNumericCellValue();
			return data;
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
					for (int i=0 ; i<=ResultTypeMap.size() ; i++)
					{		
						if (cell.getCellType() == XSSFCell.CELL_TYPE_STRING && cell.getStringCellValue().equals(ResultTypeMap.get(i)))
							{
								System.out.print(cell.getStringCellValue()+" ");
								System.out.println(getNumericData(1, cell.getRowIndex(), cell.getColumnIndex()+1));
							}
					}
					
				}
				
			 }
		
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
		
		public Map<Integer, String> addResultType() {
			int j=0;
			
			while (j<ResultType.values().length)
			{
				
				String val = ResultType.values()[j].getDesc();
				j++;
				ResultTypeMap.put(j, val);
			}
		
			return ResultTypeMap;
		}

		
		public int getRowNumber(int sheetNumber) {
			sheet1= wb.getSheetAt(sheetNumber);
			int rowCount = sheet1.getLastRowNum()-sheet1.getFirstRowNum()+1;
			return rowCount;
        }
		
		public String getStringCell(int sheetNumber,int row ,int column){
			sheet1= wb.getSheetAt(sheetNumber);
			String data =sheet1.getRow(row).getCell(column).getStringCellValue();
			return data;
		}
	
		
}
