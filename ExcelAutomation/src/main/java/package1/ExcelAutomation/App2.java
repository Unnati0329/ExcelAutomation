package package1.ExcelAutomation;

import java.util.ArrayList;

import com.spire.xls.Workbook;
import com.spire.xls.Worksheet;
import com.spire.xls.collections.AutoFiltersCollection;

public class App2 
{
	static void sheetOf(String reviewer) {
		Workbook wb = new Workbook();
    	wb.loadFromFile(".\\files\\StatusReport.xlsx");
		Worksheet sheet = wb.getWorksheets().get(0);
		AutoFiltersCollection filters = sheet.getAutoFilters();
		filters.addFilter(2,reviewer);
        filters.filter();
        String reviewerName= reviewer.substring(0, reviewer.indexOf('('));
        wb.saveToFile("Manager Access Review for "+reviewerName+".xlsx");
        System.out.println(reviewerName+"sheet made successfully"); 
	}
	
    public static void main( String[] args )
    {
    	ArrayList<String> reviewer= new ArrayList<String>();
    	reviewer.add("Daniel Willis (1000114)");
    	reviewer.add("Unnati Ahuja (1064991)");
    	for(String element:reviewer) {
			sheetOf(element);
		}
    }
}