package Base_Code;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

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

        try (Workbook workbook = new XSSFWorkbook(); FileOutputStream fileOut = new FileOutputStream("DoctorFeesInfo.xlsx")) {
            Sheet sheet = workbook.createSheet("Doctor Info");
            Row headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue("Doctor Name");
            headerRow.createCell(1).setCellValue("Doctor Fees");
            headerRow.createCell(2).setCellValue("Doctor Interested Areas");

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
                        } catch (org.openqa.selenium.NoSuchElementException e) {
                            doctorFees = "NA";
                        }

                        Row row = sheet.createRow(rowNum++);
                        if (doctorName != null && !doctorName.isEmpty()) {
                            row.createCell(0).setCellValue(doctorName);
                        } else {
                            row.createCell(0).setCellValue("NA");
                        }
                        if (doctorFees != null && !doctorFees.isEmpty()) {
                            row.createCell(1).setCellValue(doctorFees);
                        } else {
                            row.createCell(1).setCellValue("NA");
                        }
                        if (doctorInterestAreas != null && !doctorInterestAreas.isEmpty()) {
                            row.createCell(2).setCellValue(doctorInterestAreas);
                        } else {
                            row.createCell(2).setCellValue("NA");
                        }

                        System.out.println("Doctor Name: " + doctorName);
                        System.out.println("Doctor Fees: " + doctorFees);
                        System.out.println("Doctor Specialization: " + doctorInterestAreas);
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
