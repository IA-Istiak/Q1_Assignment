package dataDrivenTesting;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;

import org.apache.poi.xssf.usermodel.*;

public class DataDrivenTesting {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		XSSFWorkbook ExcelWbook = null;
		XSSFSheet ExcelSheet;
		XSSFRow Row;
		XSSFCell Cell;
		//Creating an object of File class for open file
		File excelFile = new File("D:\\ProgramFiles\\JAVA_automation\\Q1_Assignment\\Excel.xlsx");
		//Creating an object of FileInputStream to read data from file
		try {
			FileInputStream inputStream = new FileInputStream(excelFile);
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}

}
