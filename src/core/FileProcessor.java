package core;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class FileProcessor {
	
	File file;
	Sheet sheetToSort;
	ExcelSheetStorage excelStorage;
		
	public FileProcessor(){};
	
	public void setFile(File f){
		file = f;
	}
	
	public boolean fetchFile(String excelFilePath) throws IOException {
		// TODO check the excel file fetching
			FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
			Workbook workbook;
			if(excelFilePath.contains(".xlsx")){
				workbook = new XSSFWorkbook(inputStream);
			} else if (excelFilePath.contains(".xls")) {
				workbook = new HSSFWorkbook(inputStream);
			} else {
				inputStream.close();
				return false;
			}
			sheetToSort = workbook.getSheetAt(0);
			
			excelStorage = new ExcelSheetStorage();	
			
			Iterator<Row> iterator = sheetToSort.iterator();
			int rowNum = 0;
		  	while (iterator.hasNext()) {
		  		Row row = iterator.next();
		  		removeApostropheFromRow(row, rowNum);
				//go through, set everything that is after line 4 (index 3) and not in col 1 (index 0) to be a NUMERIC 
		  		if(rowNum>3){
		  			setRowCellTypes(row);
		  		}
		  		excelStorage.addRow(row);
		  		rowNum++;
		  	}
			
			workbook.close();
			inputStream.close();
			
			return true;
	}

	private void setRowCellTypes(Row row) {
		Iterator<Cell> cellIterator = row.cellIterator();
		int i=0;
		
		while (cellIterator.hasNext()) {
			Cell cell = cellIterator.next();
			
			if(cell.getCellTypeEnum() == null ){
				cell.setCellValue(new HSSFRichTextString(""));
				cell.setCellType(CellType.BLANK);
			}else if(i>0 && cell.getCellTypeEnum()==CellType.STRING){
				cell.setCellValue(Double.parseDouble(cell.getStringCellValue()));
				cell.setCellType(CellType.NUMERIC);
			}

			
			
			i++;
		}
	}

	//Removed the ' from the front of words and adds Name as a heading in column 1 
	private void removeApostropheFromRow(Row row, int rowNum){
		Iterator<Cell> cellIterator = row.cellIterator();
		
		while (cellIterator.hasNext()) {
			Cell cell = cellIterator.next();
			
			if(cell.getCellTypeEnum() == CellType.STRING && cell.getStringCellValue().indexOf("'") == 0) {
				if(cell.getStringCellValue().equals("' ") && rowNum==3){
					cell.setCellValue("Name");
				} else {
					cell.setCellValue(cell.getStringCellValue().substring(1));						
				}

			}
			
		}
	}
	
	//input: int index of the column to sort by
	public boolean  processFile(int colNum) {

		
		//logic for sorting excel sheet
		excelStorage.sort(colNum);
		
		//printing/saving excel sheet
		excelStorage.printSheet();

	  	return true;
		
	}
	
	public ArrayList<String> getSheetColumnHeaders(){
		Row headersRow = excelStorage.getColumnHeaders();
		ArrayList<String> arrayOfHeaders = new ArrayList<String>();
		
		Iterator<Cell> cellIterator = headersRow.cellIterator();

		while (cellIterator.hasNext()) {
			Cell cell = cellIterator.next();
			arrayOfHeaders.add(cell.getStringCellValue());	
		}
		
		return arrayOfHeaders;
	}

}
