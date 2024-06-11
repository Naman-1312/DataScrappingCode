package Base_Code;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashSet;
import java.util.List;
import java.util.NoSuchElementException;
import java.util.Set;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.edge.EdgeDriver;

import io.github.bonigarcia.wdm.WebDriverManager;

public class Doctor_FeesScrapping {

    public static void main(String[] args) {
        WebDriver driver = null;
        WebDriverManager.edgedriver().setup();
        driver = new EdgeDriver();
        driver.manage().window().maximize();

        // Url opening
        driver.get("https://kivihealth.com/jaipur/doctors");
        Set<String> uniqueUrls = new HashSet<>();

        try (Workbook workbook = new XSSFWorkbook(); FileOutputStream fileOut = new FileOutputStream("DoctorFeesInfo.xlsx")) {
            Sheet sheet = workbook.createSheet("Doctor Info");
            Row headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue("Doctor Name");
            headerRow.createCell(1).setCellValue("Doctor Fees");
            headerRow.createCell(2).setCellValue("Doctor Interested Areas");
            headerRow.createCell(3).setCellValue("Doctor Profile Url");

            int rowNum = 1;
            while (true) {
                try {
                    List<WebElement> list = driver.findElements(By.xpath("//div[@class='searchContainer']//div[contains(@class,'docBox')]"));
                    for (int i = 0; i < list.size(); i++) {
                        System.out.println("*******************************" + i + "***********************");
                        WebElement element = list.get(i);
                        String doctorName = element.findElement(By.xpath(".//h4")).getAttribute("innerText");
                        String doctorFees;
                        String doctorInterestAreas = element.findElement(By.xpath(".//h5[2]")).getAttribute("innerText");
                        try {
                            doctorFees = element.findElement(By.className("fee-charges")).getText();
                        } catch (NoSuchElementException e) {
                            doctorFees = "NA";
                        }

                        String imageprofileUrl;
                        try {
                            imageprofileUrl = element.findElement(By.xpath(".//img")).getAttribute("src");
                        } catch (NoSuchElementException e) {
                            imageprofileUrl = "NA";
                        }

                        Row row = sheet.createRow(rowNum++);
                        row.createCell(0).setCellValue(doctorName != null && !doctorName.isEmpty() ? doctorName : "NA");
                        row.createCell(1).setCellValue(doctorFees != null && !doctorFees.isEmpty() ? doctorFees : "NA");
                        row.createCell(2).setCellValue(doctorInterestAreas != null && !doctorInterestAreas.isEmpty() ? doctorInterestAreas : "NA");
                        row.createCell(3).setCellValue(imageprofileUrl != null && !imageprofileUrl.isEmpty() ? imageprofileUrl : "NA");

                        System.out.println("Doctor Name: " + doctorName);
                        System.out.println("Doctor Fees: " + doctorFees);
                        System.out.println("Doctor Specialization: " + doctorInterestAreas);
                        System.out.println("Doctor Profile Image: " + imageprofileUrl);
                    }
                    
                    List<WebElement> links = driver.findElements(By.tagName("a"));
                    // Loop through each link and print the href attribute if it matches a specific pattern
                    for (WebElement link : links) {
                        String url = link.getAttribute("href");
                        if (url != null && url.contains("iam")) {
                            uniqueUrls.add(url);
                        }
                    }
                    
                    driver.findElement(By.xpath("//i[contains(text(),'chevron_right')]")).click();
                } catch (Exception e) {
                    System.out.println("Exception occurred. Saving the Excel file.");
                    workbook.write(fileOut);
                    break;
                }
            }
            System.out.println("Doctor information successfully written in the Excel");
        } catch (IOException e) {
            System.out.println("Error writing to Excel: " + e);
        } finally {
            driver.quit();
        }
    }
}
