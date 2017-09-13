package core;

import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.examples.CellTypes;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;

public class ExcelSheetStorage {

	private ArrayList<Row> rows = new ArrayList<Row>();
	
	
	public ExcelSheetStorage(){	}

	public void addRow(Row r){
		rows.add(r);
	}
	

	public Row getColumnHeaders(){
		return rows.get(3);
	}


	//input: the index of the column to be sorted by (if 0, sort alphabetically, else sort numerically high->low)
	public void sort(int colNum){ 
		
		if(sheetHasSubGroups()){
			sortWithSubgroups(colNum);
		} else {
			sortWithoutSubgroups(colNum);
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
		// TODO Auto-generated method stub

		ArrayList<Pair> headerPairs = new ArrayList<Pair>();
		findHeaderPairs(headerPairs);
		//findSubgroups(colNum);
		
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

	private void sortWithoutSubgroups(int colNum) {
		int bottom = 4;
		int top = 3;
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
		boolean hasValues = true;
		for(int i=0; i<r.getLastCellNum();i++) {
			Cell c = r.getCell(i,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
			
			if(i==0 && c == null || c.getCellTypeEnum() == null || c.getCellTypeEnum() == CellType.BLANK) { //first cell is blank, that row should not move, return true
				break;
			} else if(i==0 && c.getCellTypeEnum() == CellType.STRING && r.getCell(i).getStringCellValue()!=null && isUpperCase(r.getCell(i).getStringCellValue())){ //name is all CAPS, that row should not move, return true
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
		System.out.println();
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
