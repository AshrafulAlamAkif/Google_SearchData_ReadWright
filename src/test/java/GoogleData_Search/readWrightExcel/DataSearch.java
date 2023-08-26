package GoogleData_Search.readWrightExcel;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.Test;

import io.github.bonigarcia.wdm.WebDriverManager;

public class DataSearch {

	@Test
	public void FirstSearch() throws InterruptedException, IOException  {
        WebDriverManager.chromedriver().setup();
        WebDriver driver = new ChromeDriver();

        driver.get("https://www.google.com");

        FileInputStream excelFile = new FileInputStream("F:\\Software 2\\Java and Mavan\\Excel_1.xlsx");
        Workbook workbook = new XSSFWorkbook(excelFile);
        Sheet sheet = workbook.getSheet("Saturday");

        String[] searchData = new String[10];
        for (int i = 0; i < 10; i++) {
        	searchData[i] = sheet.getRow(i + 2).getCell(2).getStringCellValue();
        }

        for (int i = 0; i < searchData.length; i++) {
            WebElement searchBox = driver.findElement(By.name("q"));
            searchBox.clear();
            searchBox.sendKeys(searchData[i]);

            Thread.sleep(1000);

            List<WebElement> suggestionList = driver.findElements(By.xpath("//div[@id='Alh6id']//li[@role='presentation']//div[@class='wM6W7d']"));

            String largestData = "";
            String shortestData = suggestionList.get(0).getText();

            for (WebElement suggestion : suggestionList) {
                String suggestionText = suggestion.getText();
                if (suggestionText.length() > largestData.length()) {
                    largestData = suggestionText;
                } else if (suggestionText.length() < shortestData.length()) {
                    shortestData = suggestionText;
                }
            }

            System.out.println("Search Data: " + searchData[i]);
            System.out.println("Largest Data: " + largestData);
            System.out.println("Shortest Data: " + shortestData);
            System.out.println("------------------------------");

            // Write the largest and smallest options to the Excel file
            Row row = sheet.getRow(i + 2); 					// Adjust index for rows starting from 2
            Cell longestOptionCell = row.createCell(3); 	// Column D (index 2)
            Cell shortestOptionCell = row.createCell(4); 	// Column E (index 3)
            longestOptionCell.setCellValue(largestData);
            shortestOptionCell.setCellValue(shortestData);
        }

        FileOutputStream outFile = new FileOutputStream("F:\\Software 2\\Java and Mavan\\Excel_1.xlsx");
        workbook.write(outFile);
        outFile.close();

        // Close the workbook and input stream
        workbook.close();
        excelFile.close();
        
        
        driver.quit();
    }
}

