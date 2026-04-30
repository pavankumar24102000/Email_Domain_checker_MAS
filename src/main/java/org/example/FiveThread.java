package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;

import java.io.*;
import java.net.URI;
import java.net.http.*;
import java.time.Duration;
import java.util.*;
import java.util.concurrent.*;
import java.util.concurrent.atomic.AtomicInteger;

public class FiveThread {

    static final List<String> SKIP_VALUES = List.of(
            "row labels", "email", "grand total", "status", ""
    );

    static class EmailResult {
        String email, status, colorLabel;
        int rowIndex;
        EmailResult(String email, String status, String colorLabel, int rowIndex) {
            this.email = email; this.status = status;
            this.colorLabel = colorLabel; this.rowIndex = rowIndex;
        }
    }

    static AtomicInteger safeCount       = new AtomicInteger(0);
    static AtomicInteger disposableCount = new AtomicInteger(0);
    static AtomicInteger unknownCount    = new AtomicInteger(0);
    static List<EmailResult> results     = Collections.synchronizedList(new ArrayList<>());

    public static void main(String[] args) throws Exception {
        String inputPath  = "C:\\Users\\pavan.k1\\Downloads\\Domain.xlsx";
        String outputPath = "C:\\Users\\pavan.k1\\Downloads\\output_scrape.xlsx";

        // 📥 Read emails from Excel
        FileInputStream fis = new FileInputStream(inputPath);
        Workbook inputWorkbook = new XSSFWorkbook(fis);
        Sheet inputSheet = inputWorkbook.getSheetAt(0);
        DataFormatter formatter = new DataFormatter();

        List<String> emails = new ArrayList<>();
        for (int i = 0; i <= inputSheet.getLastRowNum(); i++) {
            Row row = inputSheet.getRow(i);
            if (row == null) continue;
            Cell cell = row.getCell(0);
            if (cell == null) continue;
            String email = formatter.formatCellValue(cell).trim();
            if (SKIP_VALUES.contains(email.toLowerCase())) {
                System.out.println("⏭️  Skipping: " + email);
                continue;
            }
            emails.add(email);
        }
        fis.close();
        inputWorkbook.close();

        System.out.println("=============================================================");
        System.out.println("📧 Total to process : " + emails.size());
        System.out.println("🚀 Scraping verifymail.io directly — no API key needed!");
        System.out.println("=============================================================");
        System.out.printf("%-50s %-35s %-10s%n", "Email", "Status", "Color");
        System.out.println("=============================================================");

        // ✅ Shared HTTP client
        HttpClient httpClient = HttpClient.newBuilder()
                .connectTimeout(Duration.ofSeconds(15))
                .version(HttpClient.Version.HTTP_1_1)
                .followRedirects(HttpClient.Redirect.ALWAYS)
                .build();

        // 🧵 3 threads — enough to be fast without getting blocked
        int THREAD_COUNT = 1;
        ExecutorService executor = Executors.newFixedThreadPool(THREAD_COUNT);
        List<Future<?>> futures = new ArrayList<>();

        for (int i = 0; i < emails.size(); i++) {
            final String email = emails.get(i);
            final int rowIndex = i;

            futures.add(executor.submit(() -> {
                String status     = "UNKNOWN";
                String colorLabel = "🟡 YELLOW";

                try {
                    // ✅ Hit the domain page directly — same URL Selenium was using
                    String url = "https://verifymail.io/email/" + email;

                    HttpRequest request = HttpRequest.newBuilder()
                            .uri(URI.create(url))
                            .timeout(Duration.ofSeconds(15))
                            // ✅ Mimic a real browser so site doesn't block us
                            .header("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36")
                            .header("Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8")
                            .header("Accept-Language", "en-US,en;q=0.5")
                            .GET()
                            .build();

                    HttpResponse<String> response = httpClient.send(
                            request, HttpResponse.BodyHandlers.ofString()
                    );

                    if (response.statusCode() == 200) {
                        // ✅ Parse HTML with Jsoup
                        Document doc = Jsoup.parse(response.body());

                        // Target exactly: <h2 class="display-5 text-center">is safe.</h2>
                        Element h2 = doc.selectFirst("h2.display-5.text-center");

                        if (h2 != null) {
                            String rawText = h2.text().trim().toLowerCase();
                            status = h2.text().trim(); // keep original casing for Excel

                            if (rawText.contains("safe")) {
                                colorLabel = "🟢 GREEN";
                                safeCount.incrementAndGet();
                            } else if (rawText.contains("temporary") || rawText.contains("disposable")) {
                                colorLabel = "🔴 RED";
                                disposableCount.incrementAndGet();
                            } else {
                                colorLabel = "🟡 YELLOW";
                                unknownCount.incrementAndGet();
                            }
                        } else {
                            // ✅ Fallback: check the description meta tag
                            Element meta = doc.selectFirst("meta[name=description]");
                            if (meta != null) {
                                String desc = meta.attr("content").toLowerCase();
                                if (desc.contains("safe")) {
                                    status = "is safe.";
                                    colorLabel = "🟢 GREEN";
                                    safeCount.incrementAndGet();
                                } else if (desc.contains("temporary") || desc.contains("disposable")) {
                                    status = "is a temporary/disposable email.";
                                    colorLabel = "🔴 RED";
                                    disposableCount.incrementAndGet();
                                } else {
                                    status = "Unknown (meta fallback)";
                                    colorLabel = "🟡 YELLOW";
                                    unknownCount.incrementAndGet();
                                }
                            } else {
                                status = "Element not found";
                                colorLabel = "🟡 YELLOW";
                                unknownCount.incrementAndGet();
                            }
                        }

                    } else if (response.statusCode() == 429) {
                        status     = "RATE LIMITED ⚠️";
                        colorLabel = "🟡 YELLOW";
                        unknownCount.incrementAndGet();
                        Thread.sleep(5000); // back off 5s on rate limit
                    } else {
                        status     = "HTTP ERROR " + response.statusCode();
                        colorLabel = "🟡 YELLOW";
                        unknownCount.incrementAndGet();
                    }

                } catch (Exception e) {
                    status     = "FAILED: " + e.getMessage();
                    colorLabel = "🟡 YELLOW";
                    unknownCount.incrementAndGet();
                }

                results.add(new EmailResult(email, status, colorLabel, rowIndex));
                System.out.printf("%-50s %-35s [%s]%n", email, status, colorLabel);
            }));
        }

        // Wait for all threads
        for (Future<?> f : futures) {
            try { f.get(); } catch (ExecutionException e) {
                System.out.println("❌ " + e.getMessage());
            }
        }

        executor.shutdown();
        executor.awaitTermination(2, TimeUnit.HOURS);

        // Sort back to original order
        results.sort(Comparator.comparingInt(a -> a.rowIndex));

        // 📤 Write output Excel
        Workbook outputWorkbook = new XSSFWorkbook();
        Sheet outputSheet  = outputWorkbook.createSheet("Results");
        Sheet summarySheet = outputWorkbook.createSheet("Summary");

        CellStyle greenStyle = outputWorkbook.createCellStyle();
        greenStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
        greenStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        CellStyle redStyle = outputWorkbook.createCellStyle();
        redStyle.setFillForegroundColor(IndexedColors.ROSE.getIndex());
        redStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        CellStyle yellowStyle = outputWorkbook.createCellStyle();
        yellowStyle.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
        yellowStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        // Header
        Row headerRow = outputSheet.createRow(0);
        headerRow.createCell(0).setCellValue("Email");
        headerRow.createCell(1).setCellValue("Status");

        int rowIdx = 1;
        for (EmailResult result : results) {
            Row outRow = outputSheet.createRow(rowIdx++);
            outRow.createCell(0).setCellValue(result.email);
            Cell statusCell = outRow.createCell(1);
            statusCell.setCellValue(result.status);
            if (result.colorLabel.contains("GREEN"))    statusCell.setCellStyle(greenStyle);
            else if (result.colorLabel.contains("RED")) statusCell.setCellStyle(redStyle);
            else                                        statusCell.setCellStyle(yellowStyle);
        }

        // Summary sheet
        summarySheet.createRow(0).createCell(0).setCellValue("Category");
        summarySheet.getRow(0).createCell(1).setCellValue("Count");
        String[][] summary = {
                {"🟢 Safe",                    String.valueOf(safeCount.get())},
                {"🔴 Disposable / Temporary",  String.valueOf(disposableCount.get())},
                {"🟡 Unknown / Failed",         String.valueOf(unknownCount.get())},
                {"📧 Total",                   String.valueOf(results.size())}
        };
        for (int i = 0; i < summary.length; i++) {
            Row r = summarySheet.createRow(i + 1);
            r.createCell(0).setCellValue(summary[i][0]);
            r.createCell(1).setCellValue(summary[i][1]);
        }

        outputSheet.autoSizeColumn(0);
        outputSheet.autoSizeColumn(1);
        summarySheet.autoSizeColumn(0);
        summarySheet.autoSizeColumn(1);

        FileOutputStream fos = new FileOutputStream(outputPath);
        outputWorkbook.write(fos);
        fos.close();
        outputWorkbook.close();

        System.out.println("=============================================================");
        System.out.println("✅ SUMMARY");
        System.out.println("=============================================================");
        System.out.printf("  🟢 Safe                   : %d%n", safeCount.get());
        System.out.printf("  🔴 Disposable / Temporary : %d%n", disposableCount.get());
        System.out.printf("  🟡 Unknown / Failed       : %d%n", unknownCount.get());
        System.out.printf("  📧 Total Processed        : %d%n", results.size());
        System.out.println("=============================================================");
        System.out.println("✅ Output saved to: " + outputPath);
    }
}