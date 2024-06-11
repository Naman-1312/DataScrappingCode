package Base_Code;

import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.List;
import java.util.NoSuchElementException;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import io.github.bonigarcia.wdm.WebDriverManager;

public class Interested_Areas {
    public static void main(String[] args) {
        WebDriver driver = null;
        WebDriverManager.edgedriver().setup();
        driver = new EdgeDriver();
        driver.manage().window().maximize();
        
        // URL opening
        driver.get("https://kivihealth.com/jaipur/doctors");

        // Initialize Excel workbook and sheet
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Doctor Profile");
        
        // Create header row
        Row headerRow = sheet.createRow(0);
        headerRow.createCell(0).setCellValue("Doctor Name");
        headerRow.createCell(1).setCellValue("Doctor Consultancy Fees");
        headerRow.createCell(2).setCellValue("Doctor Education");
        headerRow.createCell(3).setCellValue("Doctor Specialization");
        headerRow.createCell(4).setCellValue("Doctor Experience");
        headerRow.createCell(5).setCellValue("Doctor Profile Image");
        headerRow.createCell(6).setCellValue("Interested Area 1");
        headerRow.createCell(7).setCellValue("Interested Area 2");
        headerRow.createCell(8).setCellValue("Interested Area 3");
        headerRow.createCell(9).setCellValue("Interested Area 4");
        headerRow.createCell(10).setCellValue("Interested Area 5");

        int rowNum = 1;

        try {
            while (driver.findElement(By.xpath("//i[contains(text(),'chevron_right')]")).isDisplayed()) {
                WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(30));
                List<WebElement> list = driver.findElements(By.xpath("//div[@class='searchContainer']//div[contains(@class,'docBox')]"));
                
                // Iterate through each doctor's profile
                for (WebElement doctorProfile : list) {
                    String doctorName = doctorProfile.findElement(By.xpath(".//h4")).getText();
                    String doctorEducation = doctorProfile.findElement(By.xpath(".//h5[1]")).getText();
                    String doctorSpecialization = doctorProfile.findElement(By.xpath(".//h5[2]")).getText();
                    String doctorExperience = doctorProfile.findElement(By.xpath(".//h5[3]")).getText();
                    String imageUrl = doctorProfile.findElement(By.xpath(".//img")).getAttribute("src");
                    String doctorFees;
                    try {
                      doctorFees = doctorProfile.findElement(By.className("fee-charges")).getText();
                    } catch (NoSuchElementException e) {
                        doctorFees = "NA";
                    
                    
                    
                    
                    // Extract interested areas
                    List<WebElement> interestedAreas = doctorProfile.findElements(By.xpath(".//ul/li"));
                    String[] interestedAreaArray = new String[5];
                    for (int i = 0; i < interestedAreaArray.length; i++) {
                        interestedAreaArray[i] = i < interestedAreas.size() ? interestedAreas.get(i).getText() : "NA";
                    }
                    
                    // Write data to Excel
                    Row row = sheet.createRow(rowNum++);
                    row.createCell(0).setCellValue(doctorName != null && !doctorName.isEmpty() ? doctorName : "NA");
                    row.createCell(1).setCellValue(doctorFees != null && !doctorFees.isEmpty() ? doctorFees : "NA");
                    row.createCell(2).setCellValue(doctorEducation != null && !doctorEducation.isEmpty() ? doctorEducation : "NA");
                    row.createCell(3).setCellValue(doctorSpecialization != null && !doctorSpecialization.isEmpty() ? doctorSpecialization : "NA");
                    row.createCell(4).setCellValue(doctorExperience != null && !doctorExperience.isEmpty() ? doctorExperience : "NA");
                    row.createCell(5).setCellValue(imageUrl != null && !imageUrl.isEmpty() ? imageUrl : "NA");
                    for (int i = 0; i < interestedAreaArray.length; i++) {
                        row.createCell(6 + i).setCellValue(interestedAreaArray[i]);
                    }
                }
                driver.findElement(By.xpath("//i[contains(text(),'chevron_right')]")).click();
            }
        }
        }
            catch (Exception e) {
            System.out.println("Exception occurred. Saving the Excel file.");
            e.printStackTrace();
        } finally {
            // Save the Excel file
            try (FileOutputStream fileOut = new FileOutputStream("DoctorFrontProfileInfo.xlsx")) {
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
}
