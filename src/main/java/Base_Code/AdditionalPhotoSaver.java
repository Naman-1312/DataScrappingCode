package Base_Code;

import org.apache.commons.io.FileUtils;
import org.apache.http.HttpEntity;
import org.apache.http.client.methods.CloseableHttpResponse;
import org.apache.http.client.methods.HttpGet;
import org.apache.http.impl.client.CloseableHttpClient;
import org.apache.http.impl.client.HttpClients;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

public class AdditionalPhotoSaver {

    public static void main(String[] args) {
        String excelFilePath = "C:\\Users\\naman\\Desktop\\GulbargDoctorAdditionalPhoto.xlsx";
        String downloadDir = "C:\\Users\\naman\\Desktop\\AdditionalPhotos\\";

        try (FileInputStream fis = new FileInputStream(excelFilePath);
             Workbook workbook = new XSSFWorkbook(fis);
             CloseableHttpClient httpClient = HttpClients.createDefault()) {

            Sheet sheet = workbook.getSheetAt(0);

            for (Row row : sheet) {
                if (row.getRowNum() == 0) continue; // skip header row

                String[] guids = new String[5];
                for (int i = 0; i < 5; i++) {
                    guids[i] = getCellValue(row.getCell(i));
                }

                String[] urls = new String[5];
                for (int i = 0; i < 5; i++) {
                    urls[i] = getCellValue(row.getCell(5 + i));
                }

                for (int i = 0; i < urls.length; i++) {
                    if (urls[i] != null && !urls[i].isEmpty() && !urls[i].equalsIgnoreCase("NA") && guids[i] != null && !guids[i].isEmpty()) {
                        String filePath = downloadDir + guids[i] + ".jpg";
                        String fileName = downloadImage(httpClient, urls[i], filePath);
                        if (fileName != null) {
                            row.createCell(10 + i).setCellValue(fileName);
                            System.out.println("Downloaded: " + fileName);
                        } else {
                            System.out.println("Failed to download image from URL: " + urls[i]);
                        }
                    } else {
                        System.out.println("Skipping download for row " + row.getRowNum() + " due to missing URL, NA, or GUID.");
                    }
                }
            }

            try (FileOutputStream fos = new FileOutputStream(excelFilePath)) {
                workbook.write(fos);
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static String getCellValue(Cell cell) {
        if (cell == null) return null;
        return cell.getCellType() == CellType.STRING ? cell.getStringCellValue() : null;
    }

    private static String downloadImage(CloseableHttpClient httpClient, String url, String savePath) {
        try {
            HttpGet request = new HttpGet(url);
            try (CloseableHttpResponse response = httpClient.execute(request)) {
                HttpEntity entity = response.getEntity();
                if (entity != null) {
                    try (InputStream is = entity.getContent()) {
                        File targetFile = new File(savePath);
                        FileUtils.copyInputStreamToFile(is, targetFile);
                        return targetFile.getName();
                    }
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return null;
    }
}
