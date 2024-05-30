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

        List<String> urls = readUrlsFromExcel("DoctorUrl.xlsx");
        List<DoctorData> doctorDataList = new ArrayList<>();

        for (String url : urls) {
            driver.get(url);
            try {
                System.out.println("**********************************************************");
                System.out.println("Doctor Url : " + url);
                System.out.println();
                Thread.sleep(2000); // Sleep for 2 seconds to wait for the page to load
            } catch (InterruptedException e) {
                logger.error("Thread interrupted while waiting for page to load", e);
            }
            try {
                String doctorName = getText(driver, ".//h1");
                String doctorProfileImage = getCssValue(driver, ".doctoravtar", "background-image");
                String doctorSpecialization = getText(driver, "//*[@id='doctorprofilecard']/div[1]/p");
                String doctorRating = getText(driver, "//*[@id='doctorprofilecard']/div[2]/span[1]") + " Star";
                String doctorEducation = getText(driver, "//*[@class='educationbox']/p[2]");
                String[] additionalPhotosUrls = getAdditionalPhotos(driver);
                String[][] clinicData = getClinicData(driver);

                DoctorData doctorData = new DoctorData(doctorName, doctorProfileImage, doctorSpecialization, doctorRating, doctorEducation, additionalPhotosUrls, clinicData);
                doctorDataList.add(doctorData);

                logger.info("Doctor Name: " + doctorName);
                logger.info("Doctor Profile Image: " + doctorProfileImage);
                logger.info("Doctor Specialization: " + doctorSpecialization);
                logger.info("Doctor Ratings: " + doctorRating);
                logger.info("Doctor Education: " + doctorEducation);

            } catch (Exception e) {
                logger.error("Error scraping data from URL: " + url, e);
            }
        }

        driver.quit();
        writeDataToExcel(doctorDataList, "DoctorInfo.xlsx");
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
                "Doctor Education", "Doctor Additional Photo 1", "Doctor Additional Photo 2",
                "Doctor Additional Photo 3", "Doctor Additional Photo 4", "Doctor Additional Photo 5",
                "Doctor Clinic Name 1", "Doctor Clinic Name 2", "Doctor Clinic Name 3", "Doctor Clinic Name 4",
                "Doctor Clinic Name 5", "Doctor Clinic Phone No. 1", "Doctor Clinic Phone No. 2",
                "Doctor Clinic Phone No. 3", "Doctor Clinic Phone No. 4", "Doctor Clinic Phone No. 5"
            };

            Row headerRow = sheet.createRow(0);
            for (int i = 0; i < columns.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(columns[i]);
            }

            // Creating data rows
            int rowNum = 1;
            for (DoctorData doctorData : doctorDataList) {
                Row dataRow = sheet.createRow(rowNum++);

                dataRow.createCell(0).setCellValue(doctorData.getDoctorName());
                dataRow.createCell(1).setCellValue(doctorData.getProfileImage());
                dataRow.createCell(2).setCellValue(doctorData.getSpecialization());
                dataRow.createCell(3).setCellValue(doctorData.getRating());
                dataRow.createCell(4).setCellValue(doctorData.getEducation());

                String[] imageUrl = doctorData.getAdditionalPhotos();
                for (int i = 0; i < imageUrl.length; i++) {
                    dataRow.createCell(5 + i).setCellValue(imageUrl[i]);
                }

                String[][] clinicData = doctorData.getClinicData();
                for (int i = 0; i < clinicData.length; i++) {
                    dataRow.createCell(10 + i).setCellValue(clinicData[i][0]);
                    dataRow.createCell(15 + i).setCellValue(clinicData[i][1]);
                }
            }

            // Writing to Excel file
            try (FileOutputStream fileOut = new FileOutputStream(filePath)) {
                workbook.write(fileOut);
            }

            System.out.println("Doctor information successfully written in the Excel");
        } catch (IOException e) {
            logger.error("Error writing doctor information to Excel file", e);
        }
    }
}

// Define the DoctorData class
class DoctorData {
    private String doctorName;
    private String profileImage;
    private String specialization;
    private String rating;
    private String education;
    private String[] additionalPhotos;
    private String[][] clinicData;

    public DoctorData(String doctorName, String profileImage, String specialization, String rating, String education, String[] additionalPhotos, String[][] clinicData) {
        this.doctorName = doctorName;
        this.profileImage = profileImage;
        this.specialization = specialization;
        this.rating = rating;
        this.education = education;
        this.additionalPhotos = additionalPhotos;
        this.clinicData = clinicData;
    }

    public String getDoctorName() {
        return doctorName;
    }

    public String getProfileImage() {
        return profileImage;
    }

    public String getSpecialization() {
        return specialization;
    }

    public String getRating() {
        return rating;
    }

    public String getEducation() {
        return education;
    }

    public String[] getAdditionalPhotos() {
        return additionalPhotos;
    }

    public String[][] getClinicData() {
        return clinicData;
    }
}
