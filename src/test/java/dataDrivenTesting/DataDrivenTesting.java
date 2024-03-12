package dataDrivenTesting;

import java.util.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.DayOfWeek;
import java.time.Duration;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

import org.apache.poi.xssf.usermodel.*;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import io.github.bonigarcia.wdm.WebDriverManager;

public class DataDrivenTesting {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		WebDriver driver = null;
		XSSFWorkbook ExcelWbook = null;
		XSSFSheet ExcelSheet;
		XSSFRow Row;
		XSSFCell Cell3, Cell4;
		FileInputStream inputStream = null;
		FileOutputStream out = null;

		// Creating an object of File class for open file
		File excelFile = new File("D:\\ProgramFiles\\JAVA_automation\\Q1_Assignment\\Excel.xlsx");
		// Creating an object of FileInputStream to read data from file
		try {
			inputStream = new FileInputStream(excelFile);
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		// Need know about day name of today
		LocalDateTime now = LocalDateTime.now();
		DayOfWeek dayOfWeek = now.getDayOfWeek();
		String day = dayOfWeek.toString();
		// System.out.println(dayOfWeek);
		// System.out.println(day);

		// Create object of XSSFWorkbook to handle xlsx file
		try {
			ExcelWbook = new XSSFWorkbook(inputStream);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		// to access the sheet according the day
		for (int i = 0; i <= 6; i++) {
			if (day.equalsIgnoreCase(ExcelWbook.getSheetAt(i).getSheetName())) {
				ExcelSheet = ExcelWbook.getSheetAt(i);
				// gettotal ROW
				int totalRow = ExcelSheet.getLastRowNum();
				// total number to cell in a row
				int totalCell = ExcelSheet.getRow(2).getLastCellNum();

				for (int currentRow = 2; currentRow < totalRow; currentRow++) {

					// Launch Chrome Driver
					WebDriverManager.chromedriver().setup();
					driver = new ChromeDriver();
					driver.get("https://www.google.com/");
					// Enter key word
					driver.findElement(By.id("APjFqb")).sendKeys(ExcelSheet.getRow(currentRow).getCell(2).toString());

					try {
						Thread.sleep(3000);
					} catch (InterruptedException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}

					// Wait for the Ajax suggestion box to appear
					WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(8));
					List<WebElement> suggestionOptions = wait.until(
							ExpectedConditions.presenceOfAllElementsLocatedBy(By.xpath("//ul[@role='listbox']//li")));

					List<String> myList = new ArrayList<>();

					for (int k = 0; k < suggestionOptions.size(); k++) {
						WebElement option = suggestionOptions.get(k);
						String optionText = option.getText();
						System.out.println(option.getText());

						myList.add(optionText); // Add each value to the list

					}
					Collections.sort(myList, Comparator.comparingInt(String::length));

					try {
						// setting longest option
						Row = ExcelSheet.getRow(currentRow);

						Cell3 = Row.createCell(3);

						Cell3.setCellValue(myList.get(myList.size() - 1).toString());
						// ExcelSheet.getRow(currentRow).getCell(3).setCellValue(myList.get(myList.size()
						// - 1));

						// setting shortest option
						

						Cell4 = Row.createCell(4);

						Cell4.setCellValue(myList.get(0).toString());
						out = new FileOutputStream(
								new File("D:\\ProgramFiles\\JAVA_automation\\Q1_Assignment\\Excel.xlsx"));
					} catch (FileNotFoundException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
					try {
						ExcelWbook.write(out);
					} catch (IOException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
					try {
						out.close();
					} catch (IOException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
					

				}

			}

		}
		try {
			ExcelWbook.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		

	}

}
