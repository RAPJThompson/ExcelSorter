package core;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelSheetStorage {

	private ArrayList<Row> rows = new ArrayList<Row>();


	public ExcelSheetStorage(){	}

	public void addRow(Row r){
		rows.add(r);
	}

	public Row getRow(int i){
		return rows.get(i);
	}

	public Row getColumnHeaders(){
		return rows.get(3);
	}


	//input: the index of the column to be sorted by (if 0, sort alphabetically, else sort numerically high->low)
	public void sort(int colNum){ 

		if(sheetHasSubGroups()){
			sortWithSubgroups(colNum);
		} else {
			sortWithoutSubgroups(colNum,4,3);
		}


	}

	private boolean sheetHasSubGroups() {
		Row headingsRow = rows.get(3);
		int emptyCells = 0;
		for(int i=0;i<headingsRow.getLastCellNum()-1;i++) {
			Cell nextCell = headingsRow.getCell(i);
			if(nextCell == null || nextCell.getCellTypeEnum() == null || nextCell.getCellTypeEnum() == CellType.BLANK){
				emptyCells++;
			}
		}

		return emptyCells > 0;

	}

	private void sortWithSubgroups(int colNum) {

		ArrayList<Pair> headerPairs = new ArrayList<Pair>();
		ArrayList<RowSubgroup> subGroups = new ArrayList<RowSubgroup>();
		findHeaderPairs(headerPairs);


		int realColNum = findNthHeader(colNum);
		findSubgroups(headerPairs, subGroups, realColNum);


		if(subGroups.isEmpty()) {

			sortWithoutSubgroups(realColNum, 4,3);
		} else {
			sortWithoutSubgroups(realColNum,4,3); //First sort the subgroups
			if(realColNum < rows.get(3).getLastCellNum()){
				sortWithoutSubgroups(realColNum+1,4,3); //then sort the main column
				shiftSubgroups(subGroups); //move the subgroups to the position of their totalled line
			}
		}
	}

	//This shifts the group of rows down until it finds the string of the row that it was attached to before the sort
	private void shiftSubgroups(ArrayList<RowSubgroup> subGroups) {
		for(int i = 0; i<subGroups.size();i++) {
			RowSubgroup sg = subGroups.get(i);
			int subGroupBottom = sg.getBottom();
			int subGroupTop = sg.getTop();
			String str = sg.getSorted();

			while(subGroupTop+1 < rows.size() && !rows.get(subGroupTop+1).getCell(0).getStringCellValue().equalsIgnoreCase(str)){
				shiftDownOne(subGroupBottom, subGroupTop);
				subGroupBottom++;
				subGroupTop++;
			}
		}

	}

	//helper to move a group of rows down by one
	private void shiftDownOne(int start, int end){
		Row finalRow = rows.get(end);
		for(int i=end;i>start;i--){
			rows.set(i+1, rows.get(i));
		}
		rows.set(start, finalRow);
	}

	private void findSubgroups(ArrayList<Pair> headerPairs, ArrayList<RowSubgroup> subGroups, int colNum) {
		Pair workingPair = null; //The working pair is the header to sort on and any empty columns that are attached to it
		for(int i=0;i<headerPairs.size()-1;i++){
			Pair currentPair = headerPairs.get(i);
			if (currentPair.contains(colNum)){
				workingPair = currentPair;
				break;
			}
		}

		//If there is no working pair, then there should be no subgroups
		if(workingPair == null){
			return;
		} else {
			int lowerBound=4;
			int upperBound=3;
			int ColNumWithSubGroups = workingPair.getValOne();
			while(upperBound < rows.size()-1){
				upperBound++;
				Cell cellFromSubgroupColumn = rows.get(upperBound).getCell(ColNumWithSubGroups);

				if(cellFromSubgroupColumn != null && cellFromSubgroupColumn.getCellTypeEnum() != null && cellFromSubgroupColumn.getCellTypeEnum() != CellType.BLANK){

				} else {
					subGroups.add(new RowSubgroup(lowerBound,upperBound,rows.get(upperBound).getCell(0).getStringCellValue()));
				}
			}
		}
	}

	private int findNthHeader(int colNum) {
		Row headingsRow = rows.get(3);
		int headingNum = 0;
		int LastCellNum = headingsRow.getLastCellNum();
		for(int i=0;i<headingsRow.getLastCellNum();i++) {
			Cell nextCell = headingsRow.getCell(i);
			if(nextCell != null && nextCell.getCellTypeEnum() != null && nextCell.getCellTypeEnum() != CellType.BLANK){
				headingNum++;
			}
			if(headingNum == colNum+1){
				return i;
			}
		}
		return headingNum;
	}

	private void findHeaderPairs(ArrayList<Pair> headerPairs) {
		Row headingsRow = rows.get(3);
		for(int i=0;i<headingsRow.getLastCellNum()-1;i++) {
			Cell nextCell = headingsRow.getCell(i);
			if(nextCell == null || nextCell.getCellTypeEnum() == null || nextCell.getCellTypeEnum() == CellType.BLANK){
				headerPairs.add(new Pair(i-1,i));
			}
		}

	}

	private void sortWithoutSubgroups(int colNum, int lowerBound, int upperBound) {
		int bottom = lowerBound;
		int top = upperBound;
		while(top < rows.size()-1){
			top++;
			if(isStaticRow(rows.get(top))){
				quickSortRows(rows, bottom, top-1, colNum); 		//actual sorting
				bottom = top+1;
				while(bottom < rows.size()-1 && isStaticRow(rows.get(bottom))){
					bottom++;
				}
				top = bottom;;
			}
		}
	}

	private boolean isStaticRow(Row r){
		if(r == null){
			return false;
		}
		boolean hasValues = true;
		for(int i=0; i<r.getLastCellNum();i++) {
			Cell c = r.getCell(i,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);

			if(i==0 && (c == null || c.getCellTypeEnum() == null || c.getCellTypeEnum() == CellType.BLANK)) { //first cell is blank, that row should not move, return true
				break;
			} else if(i!=0 && c != null && c.getCellTypeEnum() != CellType.BLANK) { //has values in cells other than the name, that row should move, return false
				hasValues=false;
				break;
			}

		}
		return hasValues;
	}

	public static boolean isUpperCase(String s)
	{
		for (int i=0; i<s.length(); i++)
		{
			if (!Character.isUpperCase(s.charAt(i)))
			{
				return false;
			}
		}
		return true;
	}

	//partition for quicksort
	public int partition(ArrayList<Row> arr, int low, int high, int colNum){
		Row pivot = arr.get(high);    // pivot
		int i = (low - 1);  // Index of smaller element

		for (int j = low; j <= high- 1; j++)
		{
			// If current element is smaller than or
			// equal to pivot
			if (compareRowObjects(arr.get(j),pivot, colNum) <= 0)
			{
				i++;    // increment index of smaller element
				swapRows(arr, i, j);
			}
		}
		swapRows(arr, i + 1, high);
		return (i + 1);
	}

	//Implementation of a quicksort the rows of the excel sheet, for the indexes given, on the column given (colNum is actually nth item in the row) 
	public void quickSortRows(ArrayList<Row> arr, int low, int high, int colNum){
		if (low < high)
		{
			/* pi is partitioning index, arr[p] is now
	           at right place */
			int pi = partition(arr, low, high, colNum);

			// Separately sort elements before
			// partition and after partition
			quickSortRows(arr, low, pi - 1, colNum);
			quickSortRows(arr, pi + 1, high, colNum);
		}
	}

	public void swapRows(ArrayList<Row> arr, int row1, int row2){
		Row rowStorage = arr.get(row1);
		arr.set(row1, arr.get(row2));
		arr.set(row2, rowStorage);
	}

	protected double compareRowObjects(Row row1, Row row2, int colNum){
		Cell cell1 = row1.getCell(colNum);
		Cell cell2 = row2.getCell(colNum);

		return compareCellObjects(cell1, cell2);
	}

	protected double compareCellObjects(Cell cell1, Cell cell2) {
		if(cell1 == null || cell2 == null) {
			return (double) 0;
		} else if (cell1.getCellTypeEnum() == CellType.STRING && cell2.getCellTypeEnum() == CellType.STRING){

			return (double) cell1.getStringCellValue().compareTo(cell2.getStringCellValue());
		} else if (cell1.getCellTypeEnum() == CellType.NUMERIC && cell2.getCellTypeEnum() == CellType.NUMERIC){
			return cell1.getNumericCellValue() - cell2.getNumericCellValue();
		} else {
			System.out.println("Something strange happened in the cell comparison");
			return (double) 0;

		}
	}

	public void printSheet() {
		for(int i=0;i<rows.size();i++){

			Row row = rows.get(i); 

			if(row==null) {
			} else {
				Iterator<Cell> cellIterator = row.cellIterator();

				while (cellIterator.hasNext()) {
					Cell cell = cellIterator.next();

					switch (cell.getCellTypeEnum()) { 
					case STRING:
						System.out.print(cell.getStringCellValue());
						break;
					case BOOLEAN:
						System.out.print(cell.getBooleanCellValue());
						break;
					case NUMERIC:
						System.out.print(cell.getNumericCellValue());
						break;
					default:
						break;
					}
					System.out.print(" ");
				}
			}
			System.out.println();
		}

		
	}

	public void excelSaveSheet(Workbook workbook, String outputString) {
		Workbook wb;
		if (workbook.getClass() == HSSFWorkbook.class){
			wb = new HSSFWorkbook();
		} else {
			wb = new XSSFWorkbook();
		}
			Sheet sh = wb.createSheet();
		
		int sheetWidth = rows.get(3).getLastCellNum();
		for(int RowNum=0; RowNum<rows.size(); RowNum++){
			Row row = sh.createRow(RowNum);
			for(int ColNum=0; ColNum<sheetWidth;ColNum++){
				Cell cell = row.createCell(ColNum);

				Cell rowsCellValue = rows.get(RowNum).getCell(ColNum);
				if(rowsCellValue != null && rowsCellValue.getCellTypeEnum() != null) {
					switch (rowsCellValue.getCellTypeEnum()) { 
					case STRING:
						cell.setCellValue(rowsCellValue.getStringCellValue());
						break;
					case BOOLEAN:
						cell.setCellValue(rowsCellValue.getBooleanCellValue());
						break;
					case NUMERIC:
						cell.setCellValue(rowsCellValue.getNumericCellValue());
						break;
					default:
						break;
					}

				}
			}
		}
		
		for(int i =0;i<sh.getRow(3).getLastCellNum();i++){
			sh.autoSizeColumn(i);
		}


		try {
			FileOutputStream out = new FileOutputStream(outputString);
			wb.write(out);
			out.close();
			wb.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} finally {

		}
		

	}


	/*
	  Iterator<Row> iterator = sheetToSort.iterator();
	  	while (iterator.hasNext()) {
			Row nextRow = iterator.next();
			Iterator<Cell> cellIterator = nextRow.cellIterator();

			while (cellIterator.hasNext()) {
				Cell cell = cellIterator.next();

				switch (cell.getCellTypeEnum()) { 
				case STRING:
					System.out.print(cell.getStringCellValue());
					break;
				case BOOLEAN:
					System.out.print(cell.getBooleanCellValue());
					break;
				case NUMERIC:
					System.out.print(cell.getNumericCellValue());
					break;
				default:
					break;
				}
				System.out.print(" - ");
			}
			System.out.println();
		}

	 */


}
