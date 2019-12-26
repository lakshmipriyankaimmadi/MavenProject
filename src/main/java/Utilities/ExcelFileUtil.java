package Utilities;


import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelFileUtil {

	Workbook wb;
	
public 	ExcelFileUtil() throws Throwable{
	
	FileInputStream fi=new FileInputStream("D:\\Selenium_FrameWorks\\ERP_Maven\\TestInput\\InputSheet.xlsx");
	//System.getProperty("user.dir")+"\\TestInput\\InputSheet.xlsx
	wb=WorkbookFactory.create(fi);	
  }
	public int rowcount(String sheetname){
		
		return wb.getSheet(sheetname).getLastRowNum();		
	}
	public int colcount(String sheetname){
		
		return wb.getSheet(sheetname).getRow(0).getLastCellNum();		
	}
	public String getCellData(String sheetname,int row,int column){
		
		String data="";
		if(wb.getSheet(sheetname).getRow(row).getCell(column).getCellType()==Cell.CELL_TYPE_NUMERIC)
		{			
		int celldata=(int)wb.getSheet(sheetname).getRow(row).getCell(column).getNumericCellValue();
		data=String.valueOf(celldata);
			
		}else{
		
			data=wb.getSheet(sheetname).getRow(row).getCell(column).getStringCellValue();
		}
		return data;
	}
	public void setCellData(String sheetname,int row,int column,String status) throws Throwable{
		
		Sheet ws=wb.getSheet(sheetname);
		Row rownum=ws.getRow(row);
		Cell cell=rownum.createCell(column);
		cell.setCellValue(status);
		if(status.equalsIgnoreCase("Pass")){
			
			CellStyle style=wb.createCellStyle();
			Font font=wb.createFont();
			font.setColor(IndexedColors.GREEN.getIndex());
			font.setBold(true);
			style.setFont(font);
			rownum.getCell(column).setCellStyle(style);			
		} else if(status.equalsIgnoreCase("Fail")){
			
			CellStyle style=wb.createCellStyle();
			Font font=wb.createFont();
			font.setColor(IndexedColors.RED.getIndex());
			font.setBold(true);
			style.setFont(font);
			rownum.getCell(column).setCellStyle(style);
		} else if(status.equalsIgnoreCase("Not Completed")){
			
			CellStyle style=wb.createCellStyle();
			Font font=wb.createFont();
			font.setColor(IndexedColors.BLUE.getIndex());
			font.setBold(true);
			style.setFont(font);
			rownum.getCell(column).setCellStyle(style);
			
			
		}
		FileOutputStream fo=new FileOutputStream("D:\\Selenium_FrameWorks\\ERP_Maven\\TestOutput\\HybridOutput.xlsx");
		//System.getProperty("user.dir")+"\\TestOutput\\HybridOutput.xlsx
		wb.write(fo);
		fo.close();
	}
}
