package TestPOI.ExcelTest;

import java.util.ArrayList;
import java.util.List;

// 這裡是模擬資料庫的抓到的資料
public interface DataHelp {
	// 這裡用default是為了在未來如果因為需要實作時可以不用作這個功能也不會被影響
	default List<String[]> getData(){
		List<String[]> dataList = new ArrayList<>();
		String[] row1 = new String[] {"1", "Cindy", "20"};
		String[] row2 = new String[] {"2", "Sam", "21"};
		String[] row3 = new String[] {"3", "Hory", "22"};
		dataList.add(row1);
		dataList.add(row2);
		dataList.add(row3);
		
		return dataList;
	}

}
