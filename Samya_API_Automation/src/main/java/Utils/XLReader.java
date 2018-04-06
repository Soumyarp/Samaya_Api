package Utils;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Calendar;

import org.apache.poi.*;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XLReader {
	public static String filename= System.getProperty("user.dir")+"\\src\\testdata\\TestData.xlsx";
	public FileInputStream filein = null;
	public String path;
	private XSSFWorkbook workbook = null;
	private XSSFSheet sheet = null;
	private XSSFRow row   =null;
	private XSSFCell cell = null;
	public XLReader(String path)
	{
		this.path = path;
		try {
			filein = new FileInputStream(path);
			workbook = new XSSFWorkbook(filein);
			sheet = workbook.getSheetAt(0);
			
			filein.close();
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		//workbook = new XSSFWorkbook(filein);
 catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
	}
	
	// returns the row count in a sheet
		public int getRowCount(String sheetName){
			int index = workbook.getSheetIndex(sheetName);
			if(index==-1)
				return 0;
			else{
			sheet = workbook.getSheetAt(index);
			int number=sheet.getLastRowNum()+1;
			return number;
			}
		}
		
		// ind whether sheets exists	
		public boolean isSheetExist(String sheetName){
			int index = workbook.getSheetIndex(sheetName);
			if(index==-1){
				index=workbook.getSheetIndex(sheetName.toUpperCase());
					if(index==-1)
						return false;
					else
						return true;
			}
			else
				return true;
		}
		
		// returns number of columns in a sheet	
		public int getColumnCount(String sheetName){
			// check if sheet exists
			if(!isSheetExist(sheetName))
			 return -1;
			
			sheet = workbook.getSheet(sheetName);
			row = sheet.getRow(0);
			
			if(row==null)
				return -1;
			
			return row.getLastCellNum();
			
			
			
		}
			// returns the data from a cell
			public String getCellData(String sheetName,String colName,int rowNum){
				try{
					if(rowNum <=0)
						return "";
				
				int index = workbook.getSheetIndex(sheetName);
				int col_Num=-1;
				if(index==-1)
					return "";
				
				sheet = workbook.getSheetAt(index);
				row=sheet.getRow(0);
				for(int i=0;i<row.getLastCellNum();i++){
					//System.out.println(row.getCell(i).getStringCellValue().trim());
					if(row.getCell(i).getStringCellValue().trim().equals(colName.trim()))
						col_Num=i;
				}
				if(col_Num==-1)
					return "";
				
				sheet = workbook.getSheetAt(index);
				row = sheet.getRow(rowNum-1);
				if(row==null)
					return "";
				cell = row.getCell(col_Num);
				
				if(cell==null)
					return "";
				//System.out.println(cell.getCellType());
				if(cell.getCellType()==Cell.CELL_TYPE_STRING)
					  return cell.getStringCellValue();
				else if(cell.getCellType()==Cell.CELL_TYPE_NUMERIC || cell.getCellType()==Cell.CELL_TYPE_FORMULA ){
					  
					  String cellText  = String.valueOf(cell.getNumericCellValue());
					  if (HSSFDateUtil.isCellDateFormatted(cell)) {
				           // format in form of M/D/YY
						  double d = cell.getNumericCellValue();

						  Calendar cal =Calendar.getInstance();
						  cal.setTime(HSSFDateUtil.getJavaDate(d));
				            cellText =
				             (String.valueOf(cal.get(Calendar.YEAR))).substring(2);
				           cellText = cal.get(Calendar.DAY_OF_MONTH) + "/" +
				                      cal.get(Calendar.MONTH)+1 + "/" + 
				                      cellText;
				           
				           //System.out.println(cellText);

				         }

					  
					  
					  return cellText;
				  }else if(cell.getCellType()==Cell.CELL_TYPE_BLANK)
				      return ""; 
				  else 
					  return String.valueOf(cell.getBooleanCellValue());
				
				}
				catch(Exception e){
					
					e.printStackTrace();
					return "row "+rowNum+" or column "+colName +" does not exist in xls";
				}
			}
			
			
			// returns the data from a cell
			public String getCellData(String sheetName,int colNum,int rowNum){
				try{
					if(rowNum <=0)
						return "";
				
				int index = workbook.getSheetIndex(sheetName);

				if(index==-1)
					return "";
				
			
				sheet = workbook.getSheetAt(index);
				row = sheet.getRow(rowNum-1);
				if(row==null)
					return "";
				cell = row.getCell(colNum);
				if(cell==null)
					return "";
				
			  if(cell.getCellType()==Cell.CELL_TYPE_STRING)
				  return cell.getStringCellValue();
			  else if(cell.getCellType()==Cell.CELL_TYPE_NUMERIC || cell.getCellType()==Cell.CELL_TYPE_FORMULA ){
				  
				  String cellText  = String.valueOf(cell.getNumericCellValue());
				  if (HSSFDateUtil.isCellDateFormatted(cell)) {
			           // format in form of M/D/YY
					  double d = cell.getNumericCellValue();

					  Calendar cal =Calendar.getInstance();
					  cal.setTime(HSSFDateUtil.getJavaDate(d));
			            cellText =
			             (String.valueOf(cal.get(Calendar.YEAR))).substring(2);
			           cellText = cal.get(Calendar.MONTH)+1 + "/" +
			                      cal.get(Calendar.DAY_OF_MONTH) + "/" +
			                      cellText;
			           
			          // System.out.println(cellText);

			         }

				  
				  
				  return cellText;
			  }else if(cell.getCellType()==Cell.CELL_TYPE_BLANK)
			      return "";
			  else 
				  return String.valueOf(cell.getBooleanCellValue());
				}
				catch(Exception e){
					
					e.printStackTrace();
					return "row "+rowNum+" or column "+colNum +" does not exist  in xls";
				}
			}
		
	/*	public int getCellRowNum(String sheetName,String colName,String cellValue){
			
			for(int i=0;i<=getRowCount(sheetName);i++){
		    	if(getCellData(sheetName,colName , i).equalsIgnoreCase(cellValue)){
		    		return i;
		    	}
		    }
			return -1;
			
		}*/
		public int getCellColNum(String sheetName,String cellValue){
			
			for(int i=0; i <= getRowCount(sheetName) ; i++)
	        {
				for (int j=0;j< row.getLastCellNum();j++)
				{
				
					if(getCellData(sheetName,j, i).trim().equalsIgnoreCase(cellValue))
	                return j;
				}
	        }
			return -1;
		}
	
			public int totalColNum_samedatastartswith(String sheetName,String cellValue){
			int k =-1;
			for(int i=0; i <= getRowCount(sheetName) ; i++)
	        {
				for (int j=0;j< row.getLastCellNum();j++)
				{
				
					if(getCellData(sheetName,j, i).trim().startsWith(cellValue))
	                k = k+1;
				}
	        }
			return k ;
		}
			public int firstColNum_samedatastartswith(String sheetName,String cellValue){
				int k =-1;
				for(int i=0; i <= getRowCount(sheetName) ; i++)
		        {
					for (int j=0;j< row.getLastCellNum();j++)
					{
					
						if(getCellData(sheetName,j, i).trim().startsWith(cellValue))
		                return j;
					}
		        }
				return -1 ;
			}

}
