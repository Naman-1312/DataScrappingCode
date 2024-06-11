package individualPageData;

import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.List;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import io.github.bonigarcia.wdm.WebDriverManager;

public class DoctorFeeAndInterestedArea {
    public static void main(String[] args) {
        WebDriver driver = null;
        WebDriverManager.edgedriver().setup();
        driver = new EdgeDriver();
        driver.manage().window().maximize();
        
        // URL opening
        driver.get("https://kivihealth.com/Bareilly/doctors");

        // Initialize Excel workbook and sheet
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Bareilly Doctor Profile");
        
        // Create header row
        Row headerRow = sheet.createRow(0);
        headerRow.createCell(0).setCellValue("Doctor Name");
        headerRow.createCell(1).setCellValue("Doctor Education");
        headerRow.createCell(2).setCellValue("Doctor Specialization");
        headerRow.createCell(3).setCellValue("Doctor Experience");
        headerRow.createCell(4).setCellValue("Doctor Profile Image");
        headerRow.createCell(5).setCellValue("Doctor Consultancy Fees");
        headerRow.createCell(6).setCellValue("Doctor Interested Area 1");
        headerRow.createCell(7).setCellValue("Doctor Interested Area 2");
        headerRow.createCell(8).setCellValue("Doctor Interested Area 3");
        headerRow.createCell(9).setCellValue("Doctor Interested Area 4");
        headerRow.createCell(10).setCellValue("Doctor Interested Area 5");

        int rowNum = 1;

        try {
            // Wait for the elements to load
            new WebDriverWait(driver, Duration.ofSeconds(10)).until(ExpectedConditions.presenceOfAllElementsLocatedBy(By.xpath("//div[@class='searchContainer']//div[contains(@class,'docBox')]")));

            List<WebElement> list = driver.findElements(By.xpath("//div[@class='searchContainer']//div[contains(@class,'docBox')]"));
            
            for (int i = 0; i < list.size(); i++) {
                WebElement element = list.get(i);
                String imageUrl = getElementText(element, ".//img", "src");
                String doctorName = getElementText(element, ".//h4", "innerText");
                String doctorEducation = getElementText(element, ".//h5[1]", "innerText");
                String doctorSpecialization = getElementText(element, ".//h5[2]", "innerText");
                String doctorExperience = getElementText(element, ".//h5[3]", "innerText");
                String doctorFees = getElementText(element, ".//span[contains(@class,'fee-charges')]", "innerText");

                // Extract interested areas
                List<WebElement> interestedAreas = element.findElements(By.xpath(".//ul/li"));
                String[] interestedAreaArray = new String[5];
                for (int j = 0; j < interestedAreaArray.length; j++) {
                    interestedAreaArray[j] = j < interestedAreas.size() ? interestedAreas.get(j).getText() : "NA";
                }
                
                // Write data to Excel
                Row row = sheet.createRow(rowNum++);
                row.createCell(0).setCellValue(doctorName != null && !doctorName.isEmpty() ? doctorName : "NA");
                row.createCell(1).setCellValue(doctorEducation != null && !doctorEducation.isEmpty() ? doctorEducation : "NA");
                row.createCell(2).setCellValue(doctorSpecialization != null && !doctorSpecialization.isEmpty() ? doctorSpecialization : "NA");
                row.createCell(3).setCellValue(doctorExperience != null && !doctorExperience.isEmpty() ? doctorExperience : "NA");
                row.createCell(4).setCellValue(imageUrl != null && !imageUrl.isEmpty() ? imageUrl : "NA");
                row.createCell(5).setCellValue(doctorFees != null && !doctorFees.isEmpty() ? doctorFees : "NA");
                for (int j = 0; j < interestedAreaArray.length; j++) {
                    row.createCell(6 + j).setCellValue(interestedAreaArray[j]);
                }
            }
        } catch (Exception e) {
            System.out.println("Exception occurred. Saving the Excel file.");
            e.printStackTrace();
        } finally {
            // Save the Excel file
            try (FileOutputStream fileOut = new FileOutputStream("BareillyDoctorFeeAndInterestedArea.xlsx")) {
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
            
            // Quit the driver
            driver.quit();
        }
        System.out.println("Doctor information successfully written to the Excel file");
    }
    
    private static String getElementText(WebElement element, String xpath, String attribute) {
        try {
            WebElement el = element.findElement(By.xpath(xpath));
            return attribute.equals("innerText") ? el.getText() : el.getAttribute(attribute);
        } catch (NoSuchElementException e) {
            return "NA";
        }
    }
}
