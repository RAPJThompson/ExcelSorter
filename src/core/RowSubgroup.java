package core;


public class RowSubgroup {
	private int bottom;
	private String bottomName;
	private int top;
	private String topName;
	private String sorted;
	

	
	public RowSubgroup(int b, String bName, int t, String tName, String s){
		bottom = b;
		bottomName = bName;
		top = t;
		topName = tName;
		sorted = s;
	}
	
	
	public String getBottomName() {
		return bottomName;
	}


	public void setBottomName(String bottomName) {
		this.bottomName = bottomName;
	}


	public String getTopName() {
		return topName;
	}


	public void setTopName(String topName) {
		this.topName = topName;
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
