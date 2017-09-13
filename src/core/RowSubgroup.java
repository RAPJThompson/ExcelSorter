package core;


public class RowSubgroup {
	private int bottom;
	private int top;
	private String sorted;
	

	
	public RowSubgroup(int b, int t, String s){
		bottom = b;
		top = t;
		sorted = s;
	}
	
	
	public int getBottom() {
		return bottom;
	}

	public void setBottom(int bottom) {
		this.bottom = bottom;
	}


	public int getTop() {
		return top;
	}


	public void setTop(int top) {
		this.top = top;
	}


	public String getSorted() {
		return sorted;
	}


	public void setSorted(String sorted) {
		this.sorted = sorted;
	}


	
}
