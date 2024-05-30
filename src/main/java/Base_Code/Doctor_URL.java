package Base_Code;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.edge.EdgeDriver;
import io.github.bonigarcia.wdm.WebDriverManager;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

public class Doctor_URL {
    public static void main(String[] args) {
        WebDriver driver = null;
        WebDriverManager.edgedriver().setup();
        driver = new EdgeDriver();
        driver.manage().window().maximize();
        driver.get("https://kivihealth.com/jaipur/doctors");

        // Use a Set to store only distinct URLs
        Set<String> uniqueUrls = new HashSet<>();

        // Create a new workbook and sheet
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Doctor URLs");

        addColumnNames(sheet); // To add the column name in the excel sheet!

        // Variable to keep track of the row number
        int rowNum = 1;

        while (true) {
            try {
                // Find all the anchor elements on the page
                List<WebElement> links = driver.findElements(By.tagName("a"));

                // Loop through each link and print the href attribute if it matches a specific pattern
                for (WebElement link : links) {
                    String url = link.getAttribute("href");
                    if (url != null && url.contains("iam")) {
                        uniqueUrls.add(url);
                    }
                }

                // Click the 'Next' button to go to the next page
                driver.findElement(By.xpath("//i[contains(text(),'chevron_right')]")).click();
            } catch (NoSuchElementException e) {
                break; // Exit the loop when 'chevron_right' element is not found
            }
        }

        // Write the distinct URLs to the Excel sheet
        for (String url : uniqueUrls) {
            Row row = sheet.createRow(rowNum++);
            Cell cell = row.createCell(0);
            cell.setCellValue(url);
        }

        // Write the output to an Excel file
        try (FileOutputStream fileOut = new FileOutputStream("DoctorURLs.xlsx")) {
            workbook.write(fileOut);
        } catch (IOException e) {
            e.printStackTrace();
        }

        // Close the workbook
        try {
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }

        // Close the browser
        driver.quit();
    }

    private static void addColumnNames(Sheet sheet) {
        Row row = sheet.createRow(0);
        CellStyle style = sheet.getWorkbook().createCellStyle();
        style.setFillForegroundColor(IndexedColors.ORANGE.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        Cell cell;

        cell = row.createCell(0);
        cell.setCellValue("Doctor Url");
        cell.setCellStyle(style);
    }
}

