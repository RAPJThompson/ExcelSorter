package core;

public class Pair {
	private int valOne;
	private int valTwo;
	
	public Pair(int item1, int item2) {
		valOne = item1;
		valTwo = item2;
	}

	public int getValOne() {
		return valOne;
	}

	public void setValOne(int valOne) {
		this.valOne = valOne;
	}

	public int getValTwo() {
		return valTwo;
	}

	public void setValTwo(int valTwo) {
		this.valTwo = valTwo;
	}

	public boolean contains(int colNum) {
		if (valOne == colNum || valTwo == colNum) {
			return true;
		} else {
			return false;
		}
	}
	
	
	
}

