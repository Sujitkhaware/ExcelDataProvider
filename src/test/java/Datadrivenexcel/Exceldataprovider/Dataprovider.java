package Datadrivenexcel.Exceldataprovider;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

public class Dataprovider {
	
	//For Git Refer the Excel file attached in the Excel Branch.
	//And provide the path in FileInputStream.
	
	//There is one class called DataFormatter to change the data type from any of it to string. Defined by Apachi POI.
	DataFormatter formatter = new DataFormatter();
	
	@Test(dataProvider="driverTest")
	public void testCaseData(String Greeting, String Communication, String Id) {
		System.out.println(Greeting + " , " + Communication + " , " + Id);
	}
	
	
	@DataProvider(name="driverTest")
	public Object[][] getDate() throws IOException {
		//Object is a superset of all the data types.
		//Object[][] data = {{"Hello One", "Text One", "1"}, {"Hello Two", "Text Two", "2"}, {"Hello Three", "Text Three", "3"}};
		
		//Every row of excel we have to send it in one array.
		//Get the file in a fil object
		FileInputStream fil = new FileInputStream("F:\\Working Projects\\Selenium_Java\\Interview Questions\\Part_2\\5_DataProvider.xlsx");
		//Store the fil object in the main workbook method
		XSSFWorkbook wb = new XSSFWorkbook(fil);
		//Get the sheet which you want
		XSSFSheet sheet = wb.getSheetAt(0);
		//Get the row count
		int rowCount = sheet.getPhysicalNumberOfRows();
		//Get the row which you want
		XSSFRow row = sheet.getRow(0);
		//get the last cell index
		int columnCount = row.getLastCellNum();
		
		//Define multi dimensional array
		//So first we have created the memory to our multi dimensional array
		//I need to tell the array that how many rows needs to be the part of the array.
		Object data[][] = new Object[rowCount-1][columnCount];
		
		//Why we are using -1 because, here we are starting with 0th index.
		//Overall when the loop started for the first time, 0 will go in getRow, which i will get the A2 idex row.
		//Then we have again put in a loop stating that get me the 1st rows, cells. means getCell. get me the 0th cell for 1st row.
		//We have used the concept for outer and inner loop.
		//It will go to outer loop only when the complete inner loop is done.
		//Outer loop is giving row wise and inner loop giving column wise.
		
		//So now our duty is to capture one complete outer looop row into one array.
		//Why we are storing it in a multidimensional array because if you want to integrate with TestNG data provider, we have to do like this.
		//Then only we can pass this multidimensional array back to our Test
		
		for(int i=0;i<rowCount-1;i++) {
			//I don't want th4e get row of 0th, because i will get this header value because it is lying in the 0th row.
			//So it shall get me the first row.
			//once we get that we are storing it in a row.
			row = sheet.getRow(i+1);
			for(int j=0;j<columnCount;j++) {
				//then after every row get me the cell
				//When you pass the index it shall get the cell for that perticular row.
				//Once we captured the data for each row, we shall store it in our multidimensional array.
				//We can have any type of data in the excel, so we have to make sure that we need to first conver them into String.
				//Then store it into the multidimensional array.
				XSSFCell cell = row.getCell(j);
				data[i][j] = formatter.formatCellValue(cell);
			}
		}
		return data;
		
	}
}