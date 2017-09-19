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
			prepForSortWithSubgroups(colNum);
		} else {
			sortWithoutSubgroups(colNum,4,rows.size()-1);
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

	private void prepForSortWithSubgroups(int colNum) {

		ArrayList<Pair> headerPairs = new ArrayList<Pair>();
		ArrayList<RowSubgroup> subGroups = new ArrayList<RowSubgroup>();
		findHeaderPairs(headerPairs);


		int realColNum = findNthHeader(colNum);
		findSubgroups(headerPairs, subGroups);

		
		if(subGroups.isEmpty()) {

			sortWithoutSubgroups(realColNum, 4,rows.size()-1);
		} else {
			//printSheet();
			moveSubgroupsOutOfTheWay(headerPairs, subGroups);
			//printSheet();
			findSubgroupsNewLocation(headerPairs, subGroups);
			boolean isPartOfHeaderPair = false;
			for(int i=0;i<headerPairs.size();i++){
				Pair p = headerPairs.get(i);
				if(p.contains(realColNum)){
					isPartOfHeaderPair=true;
				}
			}
			

			if(isPartOfHeaderPair){
				sortWithSubgroups(realColNum+1,subGroups, 4,3); //sort the main column, shifted by one due to the empty row accompanying it.	

			} else {
				sortWithSubgroups(realColNum,subGroups, 4,3); //sort the main group, not shifted by an empty row
			}
			//printSheet();
			shiftSubgroups(headerPairs, subGroups); //move the subgroups to the position of their totalled line
			sortBySubgroups(realColNum, headerPairs, subGroups); //Lastly sort the subgroups


			
		}
	}

	//method to move all the subgroups to the top of the list, maintaining their grouping
	private void moveSubgroupsOutOfTheWay(ArrayList<Pair> headerPairs, ArrayList<RowSubgroup> subGroups) {
		for(int i=0;i<subGroups.size();i++){
			int bottom = subGroups.get(i).getBottom();
			int top = subGroups.get(i).getTop();
			
			while(!isStaticRow(rows.get(bottom-1))){
				shiftUpOne(bottom, top);
				bottom--;
				top--;
			}
			findSubgroupsNewLocation(headerPairs, subGroups);
		}

	}

	
	private void sortBySubgroups(int realColNum, ArrayList<Pair> headerPairs, ArrayList<RowSubgroup> subGroups) {
		// 
		findSubgroupsNewLocation(headerPairs, subGroups);
		for(int i=0;i<subGroups.size();i++){
			findSubgroupsNewLocation(headerPairs, subGroups); //Find the new Locations
			RowSubgroup gr = subGroups.get(i);
			quickSortRows(rows, gr.getBottom(), gr.getTop(), realColNum); 		
		}
	}

	private void sortWithSubgroups(int realColNum, ArrayList<RowSubgroup> subGroups, int i, int j) {
		// 
		int bottom = i;
		int top = bottom;
		while(top < rows.size()-1){
			while(bottom < rows.size()-1 && isWithinSubgroup(bottom, subGroups)){
				bottom++;
			}
			if(top<bottom){
				top=bottom;
			}
			top++;
			if(isStaticRow(rows.get(top)) || isWithinSubgroup(top, subGroups)){
				quickSortRows(rows, bottom, top-1, realColNum); 		//actual sorting
				bottom = top+1;
				while(bottom < rows.size()-1 && ((isStaticRow(rows.get(bottom)) || isWithinSubgroup(bottom, subGroups)))){
					bottom++;
				}
				if(top<bottom){
					top=bottom;
				}
			}
		}
		
	}

	private boolean rangeIsWithinSubgroup(int bottom, int top, ArrayList<RowSubgroup> subGroups) {
		for(int i=bottom;i<top;i++) {
			if(isWithinSubgroup(i, subGroups)){
				return true;
			}
		}
		return false;
	}

	
	private boolean isWithinSubgroup(int index, ArrayList<RowSubgroup> subGroups) {

		for(int j=0;j<subGroups.size();j++){
			RowSubgroup gr = subGroups.get(j);
			if(index>=gr.getBottom() && index<=gr.getTop()){
				return true;
			}
		}
		return false;
	}


	//This shifts the group of rows down until it finds the string of the row that it was attached to before the sort
	private void shiftSubgroups(ArrayList<Pair> headerPairs, ArrayList<RowSubgroup> subGroups) {
		for(int i = 0; i<subGroups.size();i++) {
			findSubgroupsNewLocation(headerPairs, subGroups);
			//printSheet();
			
			RowSubgroup sg = subGroups.get(i);
			int subGroupBottom = sg.getBottom();
			int subGroupTop = sg.getTop();
			String str = sg.getSorted();
			boolean found = false; 
			if(rowToFindIsBelow(subGroupTop, str)){
				while(subGroupTop+1 < rows.size()-1 && !found){
					String nameOfNextCell = rows.get(subGroupTop+1).getCell(0).getStringCellValue();
					if(!nameOfNextCell.equalsIgnoreCase(str)){
						shiftDownOne(subGroupBottom, subGroupTop);
						subGroupBottom++;
						subGroupTop++;
					} else {
						found = true;
						//printSheet();
					}
				}
			} else if(rowToFindIsAbove(subGroupBottom, str)){
				while(subGroupBottom-1 > 4 && !found){
					String nameOfNextCell = rows.get(subGroupTop+1).getCell(0).getStringCellValue();
					if(!nameOfNextCell.equalsIgnoreCase(str)){
						shiftUpOne(subGroupBottom, subGroupTop);
						subGroupBottom--;
						subGroupTop--;
					} else {
						found = true;
						//printSheet();
					}
				}
			} else {
				System.out.println("Error, cannot find target cell.");
			}
			
			
		}

	}



	private void findSubgroupsNewLocation(ArrayList<Pair> headerPairs, ArrayList<RowSubgroup> subGroups) {
		for(int i=0;i<subGroups.size();i++){
			boolean topFound = false;
			boolean bottomFound = false;
			String bottomName = subGroups.get(i).getBottomName();
			String topName = subGroups.get(i).getTopName();

			for(int j=4; j<rows.size()-1 && (!topFound || !bottomFound);j++){
				if(rows.get(j).getCell(0).getStringCellValue().equalsIgnoreCase(bottomName)){
					subGroups.get(i).setBottom(j);
					bottomFound = true;
				} else if(rows.get(j).getCell(0).getStringCellValue().equalsIgnoreCase(topName)){
					subGroups.get(i).setTop(j);
					topFound = true;
				}
			}
		}
		
	}

	private boolean rowToFindIsAbove(int subGroupBottom, String str) {
		for(int i=subGroupBottom;i>4;i--){
			if(rows.get(i).getCell(0).getStringCellValue().equalsIgnoreCase(str)) {
				return true;
			}
		}
		return false;
	}
	
	private boolean rowToFindIsBelow(int subGroupTop, String str) {
		for(int i=subGroupTop;i<rows.size()-1;i++){
			if(rows.get(i).getCell(0).getStringCellValue().equalsIgnoreCase(str)) {
				return true;
			}
		}
		return false;
	}

	//helper to move a group of rows down by one
	private void shiftDownOne(int start, int end){
		Row finalRow = rows.get(end+1);
		for(int i=end;i>=start;i--){
			rows.set(i+1, rows.get(i));
		}
		
		rows.set(start, finalRow);
	}
	
	//helper to move a group of rows up by one
		private void shiftUpOne(int start, int end){
			Row finalRow = rows.get(start-1);
			for(int i=start;i<=end;i++){
				rows.set(i-1, rows.get(i));
			}
			rows.set(end, finalRow);
		}
		
		//method to itterate down the first column with subgroups and to add the subgroups to an arraylist
	private void findSubgroups(ArrayList<Pair> headerPairs, ArrayList<RowSubgroup> subGroups) {
		Pair workingPair = null; //The working pair is the header to sort on and any empty columns that are attached to it
		if(headerPairs.get(0)!=null) {
			workingPair = headerPairs.get(0);
		}
		//If there is no working pair, then there should be no subgroups
		if(workingPair == null){
			return;
		} else {
			int lowerBound=4;
			int upperBound=3;
			int ColNumWithSubGroups = workingPair.getValOne();
			while(upperBound < rows.size()-1 && lowerBound < rows.size()-1){
				Cell cellFromSubgroupColumn = rows.get(lowerBound).getCell(ColNumWithSubGroups);
				if(cellFromSubgroupColumn == null || cellFromSubgroupColumn.getCellTypeEnum() == null || cellFromSubgroupColumn.getCellTypeEnum() == CellType.BLANK){
					while(lowerBound <rows.size()-1){
						lowerBound++;
						cellFromSubgroupColumn = rows.get(lowerBound).getCell(ColNumWithSubGroups);
						if(cellFromSubgroupColumn != null && cellFromSubgroupColumn.getCellTypeEnum() != null && cellFromSubgroupColumn.getCellTypeEnum() != CellType.BLANK){
							upperBound = lowerBound;
							break;
						}
					}
					if(lowerBound >=rows.size()-1) {
						break;
					}
				}
				upperBound++;
				cellFromSubgroupColumn = rows.get(upperBound).getCell(ColNumWithSubGroups);

				if(cellFromSubgroupColumn == null || cellFromSubgroupColumn.getCellTypeEnum() == null || cellFromSubgroupColumn.getCellTypeEnum() == CellType.BLANK){
					cellFromSubgroupColumn = rows.get(lowerBound).getCell(ColNumWithSubGroups);
					if(cellFromSubgroupColumn != null && cellFromSubgroupColumn.getCellTypeEnum() != null && cellFromSubgroupColumn.getCellTypeEnum() != CellType.BLANK) {
						subGroups.add(new RowSubgroup(lowerBound,rows.get(lowerBound).getCell(0).getStringCellValue(), upperBound-1,rows.get(upperBound-1).getCell(0).getStringCellValue(), rows.get(upperBound).getCell(0).getStringCellValue()));
						lowerBound=upperBound;
					}
				}
			}
		}
	}

	private int findNthHeader(int colNum) {
		Row headingsRow = rows.get(3);
		int headingNum = 0;
		int LastCellNum = headingsRow.getLastCellNum();
		for(int i=0;i<LastCellNum;i++) {
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
		int top = bottom;
		while(top < upperBound){
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

			return (double) cell1.getStringCellValue().toLowerCase().compareTo(cell2.getStringCellValue().toLowerCase());
		} else if (cell1.getCellTypeEnum() == CellType.NUMERIC && cell2.getCellTypeEnum() == CellType.NUMERIC){
			return cell1.getNumericCellValue() - cell2.getNumericCellValue();
		} else if (cell1.getCellTypeEnum() == CellType.BLANK && cell2.getCellTypeEnum() == CellType.BLANK){
			return (double) 0;
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
		System.out.println();
		System.out.println("-----------------------------------------------------------------------------");
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
			// 
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
