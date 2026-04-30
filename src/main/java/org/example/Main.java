package org.example;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;

public class Main {

    public static void main(String[] args) throws IOException, InterruptedException {
        String inputPath = "C:\\Users\\pavan.k1\\Downloads\\Domain check 30 apr.xlsx";
        String outputPath = "C:\\Users\\pavan.k1\\Downloads\\output.xlsx";

        WebDriver driver = new ChromeDriver();

        // 📥 Read input file
        FileInputStream fis = new FileInputStream(inputPath);
        Workbook inputWorkbook = new XSSFWorkbook(fis);
        Sheet inputSheet = inputWorkbook.getSheetAt(0);

        // 📤 Create new output workbook
        Workbook outputWorkbook = new XSSFWorkbook();
        Sheet outputSheet = outputWorkbook.createSheet("Results");

        DataFormatter formatter = new DataFormatter();

        // 🎨 Styles for output file
        CellStyle greenStyle = outputWorkbook.createCellStyle();
        greenStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
        greenStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        CellStyle redStyle = outputWorkbook.createCellStyle();
        redStyle.setFillForegroundColor(IndexedColors.ROSE.getIndex());
        redStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        CellStyle yellowStyle = outputWorkbook.createCellStyle();
        yellowStyle.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
        yellowStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        // 🧾 Header row
        Row header = outputSheet.createRow(0);
        header.createCell(0).setCellValue("Email");
        header.createCell(1).setCellValue("Status");

        // 🖨️ Console header
        System.out.println("=============================================================");
        System.out.printf("%-50s %-30s%n", "Email", "Status");
        System.out.println("=============================================================");

        int outputRowIndex = 1;
        int safe = 0, temporary = 0, unknown = 0;

        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));

        for (int i = 1; i <= inputSheet.getLastRowNum(); i++) {

            Row row = inputSheet.getRow(i);
            if (row == null) continue;

            Cell emailCell = row.getCell(0); // column A
            if (emailCell == null) continue;

            String email = formatter.formatCellValue(emailCell).trim();
            if (email.isEmpty()) continue;

            String status = "UNKNOWN";

            try {
                driver.get("https://verifymail.io/email/" + email);

                // Wait for the status element to appear
                WebElement statusElement = wait.until(
                        ExpectedConditions.visibilityOfElementLocated(
                                By.xpath("//*[@data-aos='fade-down']/h2[1]")
                        )
                );
                status = statusElement.getText().trim();

            } catch (Exception e) {
                status = "FAILED TO LOAD";
                System.out.println("⚠️  Could not fetch status for: " + email + " → " + e.getMessage());
            }

            // 📤 Write to Excel
            Row outRow = outputSheet.createRow(outputRowIndex++);
            outRow.createCell(0).setCellValue(email);

            Cell statusCell = outRow.createCell(1);
            statusCell.setCellValue(status);

            // 🎨 Apply color styles based on status
            String statusLower = status.toLowerCase();
            String colorLabel;

            if (statusLower.contains("safe") || statusLower.contains("valid") ||
                    statusLower.contains("deliverable") || statusLower.contains("good")) {
                statusCell.setCellStyle(greenStyle);
                colorLabel = "🟢 GREEN";
                safe++;
            } else if (statusLower.contains("temporary") || statusLower.contains("disposable") ||
                    statusLower.contains("invalid") || statusLower.contains("risky") ||
                    statusLower.contains("failed")) {
                statusCell.setCellStyle(redStyle);
                colorLabel = "🔴 RED";
                temporary++;
            } else {
                statusCell.setCellStyle(yellowStyle);
                colorLabel = "🟡 YELLOW";
                unknown++;
            }

            // 🖨️ Print to console (same format as Excel)
            System.out.printf("%-50s %-30s [%s]%n", email, status, colorLabel);
        }

        // 🖨️ Summary in console
        System.out.println("=============================================================");
        System.out.println("✅ SUMMARY");
        System.out.println("=============================================================");
        System.out.printf("  🟢 Green  (Safe/Valid)       : %d%n", safe);
        System.out.printf("  🔴 Red    (Temporary/Invalid): %d%n", temporary);
        System.out.printf("  🟡 Yellow (Unknown/Other)    : %d%n", unknown);
        System.out.printf("  📧 Total Processed           : %d%n", (safe + temporary + unknown));
        System.out.println("=============================================================");

        // 🔒 Close input
        fis.close();
        inputWorkbook.close();

        // 💾 Write output file
        FileOutputStream fos = new FileOutputStream(outputPath);
        outputWorkbook.write(fos);
        fos.close();
        outputWorkbook.close();

        driver.quit();

        System.out.println("✅ Output saved to: " + outputPath);
    }
}