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
	Workbook workbook;


	public FileProcessor(){};

	public void setFile(File f){
		file = f;
	}

	public boolean fetchFile(String excelFilePath) throws IOException {
		FileInputStream inputStream = new FileInputStream(new File(excelFilePath));

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

		int lastRowNum = sheetToSort.getLastRowNum();
		int rowNum = 0;
		while (rowNum < lastRowNum+1) {
			Row row = sheetToSort.getRow(rowNum);
			if(row == null){
				Row lastRow = excelStorage.getRow(rowNum-1);
				row = sheetToSort.createRow(rowNum);
				for(int i =0; i<lastRow.getLastCellNum();i++){
					row.createCell(i, CellType.BLANK);
					
				}
			}
			removeApostropheFromRow(row, rowNum);

			//go through, set everything that is after line 4 (index 3) and not in col 1 (index 0) to be a NUMERIC 
			if(rowNum>3){
				setRowCellTypes(row);
			}
			excelStorage.addRow(row);
			rowNum++;
		}


		inputStream.close();

		return true;
	}

	private void setRowCellTypes(Row row) {
		if(row == null) {
			return;
		}
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
		if(row == null) {
			return;
		}
		int lastCellNum = row.getLastCellNum();
		int currCell = 0;
		while (currCell<lastCellNum) {
			Cell cell = row.getCell(currCell);

			if (cell == null || cell.getCellTypeEnum() == null){
				
			}else if(cell.getCellTypeEnum() == CellType.STRING && cell.getStringCellValue().indexOf("'") == 0) {
				if(cell.getStringCellValue().equals("' ") && rowNum==3){
					cell.setCellValue("Name");
				} else {
					cell.setCellValue(cell.getStringCellValue().substring(1));						
				}

			}
			currCell++;

		}
	}

	//input: int index of the column to sort by
	public boolean  processFile(int colNum, String outputString) {


		//logic for sorting excel sheet
		excelStorage.sort(colNum);

		//printing/saving excel sheet
		//excelStorage.printSheet();
		
		
		try {
			excelStorage.excelSaveSheet(workbook, outputString);

			workbook.close();
		} catch (IOException e) {
			
			e.printStackTrace();
		}
		
		return true;

	}

	public ArrayList<String> getSheetColumnHeaders(){
		Row headersRow = excelStorage.getColumnHeaders();
		ArrayList<String> arrayOfHeaders = new ArrayList<String>();

		Iterator<Cell> cellIterator = headersRow.cellIterator();

		while (cellIterator.hasNext()) {
			Cell cell = cellIterator.next();
			String headerTitle = cell.getStringCellValue().replaceAll("/", "-");
			arrayOfHeaders.add(headerTitle);	
		}

		return arrayOfHeaders;
	}

}
