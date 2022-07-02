package package1.ExcelAutomation;

import java.awt.Color;

import com.spire.xls.Workbook;
import com.spire.xls.Worksheet;
import com.spire.xls.collections.AutoFiltersCollection;

public class TestApp 
{
	
	static void singleSheetFilter() {
//create new excel workbook
    	Workbook wb = new Workbook();
//import all the data of all the sheets from given excel file
    	wb.loadFromFile(".\\files\\StatusReport.xlsx");
    	Worksheet sheet = wb.getWorksheets().get(0);
//define filters class
    	AutoFiltersCollection filters = sheet.getAutoFilters();
//add more than one filter
    	filters.addFilter(2, "Unnati Ahuja (1064991)");
    	filters.addFilter(2, "Abhishek Pandey (1066483)");
//it'll apply all the mentioned filters
    	filters.filter();
//save excel sheet with given name
    	wb.saveToFile("testing.xlsx");
    	System.out.println("done"); 
	}
	
	static void multipleSheetsFilter() {
//create new excel workbook
    	Workbook wb = new Workbook();
//import all the data of all the sheets from given excel file
    	wb.loadFromFile(".\\files\\StatusReport.xlsx");
//define sheet1/sheet2 to apply the filters accordingly
//for eg, from sheet 1, i want to filter out Unnati and from sheet 2, i want to filter put Abhishek
    	Worksheet sheet1 = wb.getWorksheets().get(0);
    	Worksheet sheet2 = wb.getWorksheets().get(1);
//define filters class
    	AutoFiltersCollection filters1 = sheet1.getAutoFilters();
    	AutoFiltersCollection filters2 = sheet2.getAutoFilters();
//add more than one filter
    	filters1.addFilter(2, "Unnati Ahuja (1064991)");
    	//filters1.addFilter();
    	filters2.addFilter(2, "Abhishek Pandey (1066483)");
//it'll apply all the mentioned filters
    	filters1.filter();
    	filters2.filter();
//save excel sheet with given name
    	wb.saveToFile("testing.xlsx");
    	System.out.println("done"); 
	}
	
	static void formatCells() {
		   Workbook wb = new Workbook();
		   wb.loadFromFile(".\\files\\StatusReport.xlsx");
		   Worksheet sheet = wb.getWorksheets().get(0);
		   AutoFiltersCollection filters = sheet.getAutoFilters();
		   filters.addFilter(2, "Unnati Ahuja (1064991)");
		   filters.filter();
		   wb.saveToFile("testing.xlsx");
		   System.out.println("done"); 
			}
	
	public static void main( String[] args )
    {
		formatCells();
    }
}
