package Base_Code;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import io.github.bonigarcia.wdm.WebDriverManager;
import org.apache.log4j.Logger;
import org.apache.log4j.PropertyConfigurator;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class Doctor_Information {

    private static final Logger logger = Logger.getLogger(Doctor_Information.class);

    public static void main(String[] args) {
        // Configure log4j for logging
        PropertyConfigurator.configure("src/main/resources/log4j.properties");

        WebDriverManager.edgedriver().setup();
        WebDriver driver = new EdgeDriver();
        driver.manage().window().maximize();

        List<String> urls = readUrlsFromExcel("DoctorURLs.xlsx");
        List<DoctorData> doctorDataList = new ArrayList<>();

        for (String url : urls) {
            driver.get(url);
            try {
//                System.out.println("**********************************************************");
                System.out.println("Doctor Url : " + url);
//                System.out.println();
                Thread.sleep(2000); // Sleep for 2 seconds to wait for the page to load
            } catch (InterruptedException e) {
                logger.error("Thread interrupted while waiting for page to load", e);
            }
            try {
                String doctorName = getText(driver, ".//h1");
                String doctorProfileImage = getCssValue(driver, ".doctoravtar", "background-image");
                String doctorSpecialization = getText(driver, "//*[@id='doctorprofilecard']/div[1]/p");
                String doctorRating = getText(driver, "//*[@id='doctorprofilecard']/div[2]/span[1]");
                String doctorEducation = getText(driver, "//*[@class='educationbox']/p[2]");
                String[] additionalPhotosUrls = getAdditionalPhotos(driver);
                String[][] clinicData = getClinicData(driver);
                String[][] clinicAdditionalData = getClinicAdditionalData(driver);

                DoctorData doctorData = new DoctorData(doctorName, doctorProfileImage, doctorSpecialization, doctorRating, doctorEducation, additionalPhotosUrls, clinicData, clinicAdditionalData);
                doctorDataList.add(doctorData);

                logger.info("Doctor Name: " + doctorName);
                logger.info("Doctor Profile Image: " + doctorProfileImage);
                logger.info("Doctor Specialization: " + doctorSpecialization);
                logger.info("Doctor Ratings: " + doctorRating);
                logger.info("Doctor Education: " + doctorEducation);

            } catch (Exception e) {
                logger.error("Error scraping data from URL: " + url, e);
                writeDataToExcel(doctorDataList, "ErrorData.xlsx");
            }
        }

        driver.quit();
        writeDataToExcel(doctorDataList, "BellaryDoctorInfo.xlsx");
    }

    private static List<String> readUrlsFromExcel(String filePath) {
        List<String> urls = new ArrayList<>();
        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheetAt(0);
            Iterator<Row> rowIterator = sheet.iterator();

            // Skip the header row
            if (rowIterator.hasNext()) {
                rowIterator.next();
            }

            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                Cell cell = row.getCell(0);
                if (cell != null) {
                    urls.add(cell.getStringCellValue());
                }
            }
        } catch (IOException e) {
            logger.error("Error reading URLs from Excel file", e);
        }
        return urls;
    }

    private static String getText(WebDriver driver, String xpath) {
        try {
            return driver.findElement(By.xpath(xpath)).getText();
        } catch (Exception e) {
            logger.error("Error fetching text from: " + xpath, e);
            return "NA";
        }
    }

    private static String getCssValue(WebDriver driver, String cssSelector, String propertyName) {
        try {
            String cssValue = driver.findElement(By.cssSelector(cssSelector)).getCssValue(propertyName);
            if (cssValue.startsWith("url(\"") && cssValue.endsWith("\")")) {
                return cssValue.substring(5, cssValue.length() - 2);
            } else if (cssValue.startsWith("url(") && cssValue.endsWith(")")) {
                return cssValue.substring(4, cssValue.length() - 1);
            }
            return cssValue;
        } catch (Exception e) {
            logger.error("Error fetching CSS value from: " + cssSelector, e);
            return "NA";
        }
    }

    private static String[] getAdditionalPhotos(WebDriver driver) {
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
        List<WebElement> imageThumbnails = driver.findElements(By.cssSelector("#doctorphotos .image-thumbnail"));
        String[] additionalPhotosUrls = new String[5];
        for (int i = 0; i < 5; i++) {
            if (i < imageThumbnails.size()) {
                additionalPhotosUrls[i] = extractImageUrl(imageThumbnails.get(i).getAttribute("style"));
            } else {
                additionalPhotosUrls[i] = "NA";
            }
        }
        return additionalPhotosUrls;
    }

    private static String[][] getClinicData(WebDriver driver) {
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(30));
        List<WebElement> clinicNames = wait.until(ExpectedConditions.visibilityOfAllElementsLocatedBy(By.xpath("//*[@id='doctorclinics']/div/p[1]")));
        List<WebElement> phoneNumbers = wait.until(ExpectedConditions.visibilityOfAllElementsLocatedBy(By.xpath("//*[@id='doctorclinics']/div/div[2]/a[1]")));

        String[][] clinicData = new String[5][2];
        for (int i = 0; i < 5; i++) {
            clinicData[i][0] = i < clinicNames.size() ? clinicNames.get(i).getText() : "NA";
            clinicData[i][1] = i < phoneNumbers.size() ? phoneNumbers.get(i).getAttribute("href") : "NA";
        }
        return clinicData;
    }

    private static String[][] getClinicAdditionalData(WebDriver driver) {
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(30));
        wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("doctorclinics")));
        List<WebElement> clinicAddress = driver.findElements(By.xpath("//*[@id='doctorclinics']/div/p[2]"));
        List<WebElement> clinicTimming = driver.findElements(By.xpath("//*[@id='doctorclinics']/div/div/div"));

        String[][] clinicAdditionalData = new String[5][2];
        for (int i = 0; i < 5; i++) {
            clinicAdditionalData[i][0] = i < clinicAddress.size() ? clinicAddress.get(i).getText() : "NA";

            if (i < clinicTimming.size()) {
                WebElement parentDiv = clinicTimming.get(i);
                List<WebElement> pTags = parentDiv.findElements(By.tagName("p"));
                StringBuilder timings = new StringBuilder();
                for (WebElement pTag : pTags) {
                    timings.append(pTag.getText()).append(", ");
                }
                clinicAdditionalData[i][1] = timings.toString().trim();
            } else {
                clinicAdditionalData[i][1] = "NA";
            }
        }
        
        return clinicAdditionalData;
    }

    private static String extractImageUrl(String styleAttribute) {
        String imageUrl = "";
        if (styleAttribute != null && styleAttribute.contains("background-image: url(")) {
            int startIndex = styleAttribute.indexOf("background-image: url(") + "background-image: url(".length();
            int endIndex = styleAttribute.indexOf(")", startIndex);
            if (endIndex > startIndex) {
                imageUrl = styleAttribute.substring(startIndex, endIndex).replace("\"", "");
            }
        }
        return imageUrl;
    }

    private static void writeDataToExcel(List<DoctorData> doctorDataList, String filePath) {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Doctor Info");

            // Creating header row
            String[] columns = {
                "Doctor Name", "Doctor Profile Image", "Doctor Specialization and Experience", "Doctor Rating",
                "Doctor Education", "ServerImage1", "ServerImage2",
                "ServerImage3", "ServerImage4", "ServerImage5",
                "Doctor Clinic Name 1", "Doctor Clinic Name 2", "Doctor Clinic Name 3", "Doctor Clinic Name 4",
                "Doctor Clinic Name 5", "Doctor Clinic Phone No. 1", "Doctor Clinic Phone No. 2",
                "Doctor Clinic Phone No. 3", "Doctor Clinic Phone No. 4", "Doctor Clinic Phone No. 5","Clinic Address 1",
                "Clinic Address 2", "Clinic Address 3", "Clinic Address 4", "Clinic Address 5", "Clinic Timings 1",
                "Clinic Timings 2", "Clinic Timings 3", "Clinic Timings 4", "Clinic Timings 5"
            };
            Row headerRow = sheet.createRow(0);
            for (int i = 0; i < columns.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(columns[i]);
            }

            // Creating data rows
            int rowNum = 1;
            for (DoctorData doctorData : doctorDataList) {
                Row row = sheet.createRow(rowNum++);
                row.createCell(0).setCellValue(doctorData.getDoctorName());
                row.createCell(1).setCellValue(doctorData.getDoctorProfileImage());
                row.createCell(2).setCellValue(doctorData.getDoctorSpecializationAndExperience());
                row.createCell(3).setCellValue(doctorData.getDoctorRating());
                row.createCell(4).setCellValue(doctorData.getDoctorEducation());

                String[] additionalPhotos = doctorData.getAdditionalPhotos();
                for (int i = 0; i < additionalPhotos.length; i++) {
                    row.createCell(5 + i).setCellValue(additionalPhotos[i]);
                }

                String[][] clinicData = doctorData.getClinicData();
                for (int i = 0; i < clinicData.length; i++) {
                    row.createCell(10 + i).setCellValue(clinicData[i][0]);
                    row.createCell(15 + i).setCellValue(clinicData[i][1]);
                }
                
                String[][] clinicAdditionalData = doctorData.getClinicAdditionalData();
                for (int i = 0; i < clinicAdditionalData.length; i++) {
                    row.createCell(20 + i).setCellValue(clinicAdditionalData[i][0]);
                    row.createCell(25 + i).setCellValue(clinicAdditionalData[i][1]);
                }
            }

            // Auto-size all columns
            for (int i = 0; i < columns.length; i++) {
                sheet.autoSizeColumn(i);
            }

            // Write the output to a file
            try (FileOutputStream fos = new FileOutputStream(filePath)) {
                workbook.write(fos);
            }
        } catch (IOException e) {
            logger.error("Error writing data to Excel file", e);
        }
    }

    private static class DoctorData {
        private final String doctorName;
        private final String doctorProfileImage;
        private final String doctorSpecializationAndExperience;
        private final String doctorRating;
        private final String doctorEducation;
        private final String[] additionalPhotos;
        private final String[][] clinicData;
        private final String[][] clinicAdditionalData;

        public DoctorData(String doctorName, String doctorProfileImage, String doctorSpecializationAndExperience, String doctorRating, String doctorEducation, String[] additionalPhotos, String[][] clinicData, String[][] clinicAdditionalData) {
            this.doctorName = doctorName;
            this.doctorProfileImage = doctorProfileImage;
            this.doctorSpecializationAndExperience = doctorSpecializationAndExperience;
            this.doctorRating = doctorRating;
            this.doctorEducation = doctorEducation;
            this.additionalPhotos = additionalPhotos;
            this.clinicData = clinicData;
            this.clinicAdditionalData = clinicAdditionalData;
        }

        public String getDoctorName() {
            return doctorName;
        }

        public String getDoctorProfileImage() {
            return doctorProfileImage;
        }

        public String getDoctorSpecializationAndExperience() {
            return doctorSpecializationAndExperience;
        }

        public String getDoctorRating() {
            return doctorRating;
        }

        public String getDoctorEducation() {
            return doctorEducation;
        }

        public String[] getAdditionalPhotos() {
            return additionalPhotos;
        }

        public String[][] getClinicData() {
            return clinicData;
        }
        
        public String[][] getClinicAdditionalData() {
            return clinicAdditionalData;
        }
    }
}
