//copy and paste the excel sheet in files folder from which you want to filter out the data
//open the sheet->do the required formating
//save the file with StatusReport name in xlsx format
//refresh the code (File->Refresh)
//in line 32 of this code, give the name for whom you want to create the sheet

package package1.ExcelAutomation;

import java.util.ArrayList;

import org.apache.poi.ss.usermodel.*;

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
        
        Workbook wb1 = sheet.getBook();
        String reviewerName= reviewer.substring(0, reviewer.indexOf('(')); 
        wb1.saveToFile("Manager Access Review for "+reviewerName+".xlsx");
        System.out.println(reviewerName+"sheet made successfully"); 
	}
	
    public static void main( String[] args )
    {
    	ArrayList<String> reviewer= new ArrayList<String>();
    	reviewer.add("Savo Todorovic (1002350)");
    	for(String element:reviewer) {
			sheetOf(element);
		}
    }
}