package Base_Code;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;

public class Image_Saver {
    public static void main(String[] args) {
        String excelFilePath = "D:\\Data_Scraping-Selenium\\target\\Doctor Details.xlsx"; 

        try (FileInputStream fis = new FileInputStream(excelFilePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0); // Assuming the first sheet

            for (Row row : sheet) {
                Cell nameCell = row.getCell(0); // Assuming the doctor name is in the first column
                Cell urlCell = row.getCell(1); // Assuming the image URL is in the second column
               // Cell doctorGuidCell = row.getCell(2); // Assuming the Doctor Image is in the third column
                
                if (nameCell != null && urlCell != null ) {
                    String doctorName = nameCell.getStringCellValue();
                    String imageUrl = urlCell.getStringCellValue();
//                  String doctorGuid = doctorGuidCell.getStringCellValue();  

                    // Replace spaces with underscores in the doctor name for the file name
                    String fileName = doctorName.replaceAll("\\s+", "_") + "_doctorprofile.jpg";
                    String downloadPath = "C:\\Users\\naman\\Desktop\\DoctorImage\\ " + fileName;

                    downloadImage(imageUrl, downloadPath);
                }
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void downloadImage(String imageUrl, String downloadPath) {
        try {
            URL url = new URL(imageUrl);
            try (InputStream in = url.openStream();
                 FileOutputStream out = new FileOutputStream(downloadPath)) {
                byte[] buffer = new byte[1024];
                int bytesRead;
                while ((bytesRead = in.read(buffer)) != -1) {
                    out.write(buffer, 0, bytesRead);
                }
                System.out.println("Image downloaded successfully: " + downloadPath);
            } catch (IOException e) {
                e.printStackTrace();
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}