package individualPageData;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.edge.EdgeDriver;
import io.github.bonigarcia.wdm.WebDriverManager;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

public class DoctorUrl {
    public static void main(String[] args) {
        WebDriver driver = null;
        WebDriverManager.edgedriver().setup();
        driver = new EdgeDriver();
        driver.manage().window().maximize();
        driver.get("https://kivihealth.com/Bellary/doctors");

        // Use a Set to store only distinct URLs
        Set<String> uniqueUrls = new HashSet<>();

        // Find all the anchor elements on the page
        List<WebElement> links = driver.findElements(By.tagName("a"));

        // Loop through each link and add the href attribute to uniqueUrls if it matches a specific pattern
        for (WebElement link : links) {
            String url = link.getAttribute("href");
            if (url != null && url.contains("iam")) {
                uniqueUrls.add(url);
            }
        }

        // Create a new workbook and sheet
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("BellaryDoctorURLs");

        addColumnNames(sheet); // To add the column name in the excel sheet!

        // Write the distinct URLs to the Excel sheet
        int rowNum = 1;
        for (String url : uniqueUrls) {
            Row row = sheet.createRow(rowNum++);
            Cell cell = row.createCell(0);
            cell.setCellValue(url);
        }

        // Write the output to an Excel file
        try (FileOutputStream fileOut = new FileOutputStream("BellaryDoctorURLs.xlsx")) {
            workbook.write(fileOut);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            // Close the workbook
            try {
                workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }

            // Close the browser
            if (driver != null) {
                driver.quit();
            }
        }
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
