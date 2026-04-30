package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONObject;

import java.io.*;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.*;
import java.util.concurrent.atomic.AtomicInteger;

public class Myclass {

    private static final String API_KEY     = "032b5d5a8530473c93e15e801ca4c077";
    private static final String API_BASE   = "https://verifymail.io/api/";
    private static final int    THREADS    = 5;    // low to avoid 429s
    private static final int    MAX_RETRY  = 5;    // retries per email
    private static final long   RETRY_BASE = 2000; // base backoff ms, doubles each retry

    // Thread-safe counter for progress tracking
    private static final AtomicInteger processed = new AtomicInteger(0);
    private static int totalEmails = 0;

    public static void main(String[] args) throws Exception {

        String inputPath  = "C:\\Users\\pavan.k1\\Downloads\\Domain.xlsx";
        String outputPath = "C:\\Users\\pavan.k1\\Downloads\\output.xlsx";

        // ── Read all emails ──────────────────────────────────────────────────────
        List<String> emails = readEmails(inputPath);
        totalEmails = emails.size();
        System.out.println("✅ Found " + totalEmails + " email(s). Starting " + THREADS + " threads...\n");

        // ── Submit all tasks to thread pool ──────────────────────────────────────
        ExecutorService executor = Executors.newFixedThreadPool(THREADS);
        List<Future<String[]>> futures = new ArrayList<>();

        for (String email : emails) {
            futures.add(executor.submit(() -> verifyEmail(email)));
        }

        executor.shutdown();
        executor.awaitTermination(10, TimeUnit.MINUTES);

        // ── Collect results in original order ────────────────────────────────────
        List<String[]> results = new ArrayList<>();
        for (Future<String[]> future : futures) {
            try {
                results.add(future.get());
            } catch (Exception e) {
                results.add(new String[]{"ERROR", "UNKNOWN", "Failed to get result"});
            }
        }

        // ── Write Excel ──────────────────────────────────────────────────────────
        writeExcel(results, outputPath);
        System.out.println("\n✅ Done! Results saved to: " + outputPath);
    }

    // ── Read emails from column A (skip header) ──────────────────────────────────
    private static List<String> readEmails(String filePath) throws Exception {
        List<String> emails = new ArrayList<>();
        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook wb = new XSSFWorkbook(fis)) {

            Sheet sheet = wb.getSheetAt(0);
            DataFormatter formatter = new DataFormatter();
            boolean firstRow = true;

            for (Row row : sheet) {
                if (firstRow) { firstRow = false; continue; }
                Cell cell = row.getCell(0);
                if (cell != null) {
                    String val = formatter.formatCellValue(cell).trim();
                    if (!val.isEmpty()) emails.add(val);
                }
            }
        }
        return emails;
    }

    // ── Verify a single email via API with retry + exponential backoff ───────────
    private static String[] verifyEmail(String email) {
        int attempt = 0;

        while (attempt < MAX_RETRY) {
            try {
                URL url = new URL(API_BASE + email + "?key=" + API_KEY);
                HttpURLConnection conn = (HttpURLConnection) url.openConnection();
                conn.setRequestMethod("GET");
                conn.setConnectTimeout(10_000);
                conn.setReadTimeout(10_000);

                int httpStatus = conn.getResponseCode();

                // 429 = rate limited → wait and retry
                if (httpStatus == 429) {
                    attempt++;
                    long wait = RETRY_BASE * (long) Math.pow(2, attempt); // 4s, 8s, 16s...
                    System.out.printf("⚠️  429 for %-40s retry %d/%d in %ds%n",
                            email, attempt, MAX_RETRY, wait / 1000);
                    Thread.sleep(wait);
                    continue;
                }

                if (httpStatus != 200) {
                    int done = processed.incrementAndGet();
                    System.out.printf("[%d/%d] %-40s → HTTP %d%n", done, totalEmails, email, httpStatus);
                    return new String[]{email, "UNKNOWN", "-", "-", "-", "-", "-", "HTTP " + httpStatus};
                }

                StringBuilder sb = new StringBuilder();
                try (BufferedReader br = new BufferedReader(
                        new InputStreamReader(conn.getInputStream()))) {
                    String line;
                    while ((line = br.readLine()) != null) sb.append(line);
                }

                JSONObject json = new JSONObject(sb.toString());
                boolean disposable = json.optBoolean("disposable", false);
                boolean block      = json.optBoolean("block",      false);
                boolean privacy    = json.optBoolean("privacy",    false);
                boolean mx         = json.optBoolean("mx",         false);
                String  domain     = json.optString("domain",         "-");
                String  provider   = json.optString("email_provider", "-");

                String status;
                if (disposable || block) status = "Disposable - Temporary";
                else if (privacy)        status = "Privacy - Alias";
                else                     status = "Safe";

                int done = processed.incrementAndGet();
                System.out.printf("[%d/%d] %-40s → %s%n", done, totalEmails, email, status);

                return new String[]{email, status, domain, provider,
                        String.valueOf(mx), String.valueOf(disposable),
                        String.valueOf(block), String.valueOf(privacy)};

            } catch (InterruptedException ie) {
                Thread.currentThread().interrupt();
                break;
            } catch (Exception e) {
                attempt++;
                System.out.printf("⚠️  Exception for %-40s (%s) retry %d/%d%n",
                        email, e.getMessage(), attempt, MAX_RETRY);
            }
        }

        // All retries exhausted
        int done = processed.incrementAndGet();
        System.out.printf("[%d/%d] %-40s → FAILED after %d retries%n",
                done, totalEmails, email, MAX_RETRY);
        return new String[]{email, "ERROR", "-", "-", "-", "-", "-", "Failed after " + MAX_RETRY + " retries"};
    }

    // ── Write results to Excel ───────────────────────────────────────────────────
    private static void writeExcel(List<String[]> results, String outputPath) throws Exception {

        try (XSSFWorkbook wb = new XSSFWorkbook()) {

            Sheet sheet = wb.createSheet("Results");

            // ── Styles ───────────────────────────────────────────────────────────
            CellStyle headerStyle = wb.createCellStyle();
            headerStyle.setFillForegroundColor(IndexedColors.DARK_BLUE.getIndex());
            headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            headerStyle.setAlignment(HorizontalAlignment.CENTER);
            headerStyle.setBorderBottom(BorderStyle.THIN);
            Font headerFont = wb.createFont();
            headerFont.setBold(true);
            headerFont.setColor(IndexedColors.WHITE.getIndex());
            headerFont.setFontName("Arial");
            headerFont.setFontHeightInPoints((short) 11);
            headerStyle.setFont(headerFont);

            CellStyle greenStyle  = colorStyle(wb, IndexedColors.LIGHT_GREEN);
            CellStyle redStyle    = colorStyle(wb, IndexedColors.ROSE);
            CellStyle orangeStyle = colorStyle(wb, IndexedColors.LIGHT_ORANGE);
            CellStyle grayStyle   = colorStyle(wb, IndexedColors.GREY_25_PERCENT);

            // ── Header row ───────────────────────────────────────────────────────
            String[] headers = {
                    "Email", "Status", "Domain", "Email Provider",
                    "MX Valid", "Disposable", "Blocked", "Privacy / Alias"
            };
            Row headerRow = sheet.createRow(0);
            for (int i = 0; i < headers.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(headers[i]);
                cell.setCellStyle(headerStyle);
            }
            headerRow.setHeightInPoints(22);

            // ── Data rows ────────────────────────────────────────────────────────
            for (int i = 0; i < results.size(); i++) {
                String[] data = results.get(i);
                Row row = sheet.createRow(i + 1);

                CellStyle rowStyle = switch (data[1]) {
                    case "Safe"                  -> greenStyle;
                    case "Disposable - Temporary"-> redStyle;
                    case "Privacy - Alias"       -> orangeStyle;
                    default                      -> grayStyle;
                };

                for (int c = 0; c < data.length; c++) {
                    Cell cell = row.createCell(c);
                    cell.setCellValue(data[c]);
                    cell.setCellStyle(rowStyle);
                }
            }

            // ── Column widths ────────────────────────────────────────────────────
            int[] widths = {10000, 6000, 8000, 7000, 3500, 4000, 3500, 4500};
            for (int i = 0; i < widths.length; i++)
                sheet.setColumnWidth(i, widths[i]);

            // ── Summary row at bottom ────────────────────────────────────────────
            int lastRow = results.size() + 2;
            long safe       = results.stream().filter(r -> "Safe".equals(r[1])).count();
            long disposable = results.stream().filter(r -> "Disposable - Temporary".equals(r[1])).count();
            long privacy    = results.stream().filter(r -> "Privacy - Alias".equals(r[1])).count();
            long error      = results.stream().filter(r -> "ERROR".equals(r[1])).count();

            addSummaryRow(sheet, wb, lastRow,     "✅ Safe",                safe,       greenStyle);
            addSummaryRow(sheet, wb, lastRow + 1, "🚫 Disposable/Temporary",disposable, redStyle);
            addSummaryRow(sheet, wb, lastRow + 2, "🔒 Privacy/Alias",       privacy,    orangeStyle);
            if (error > 0)
                addSummaryRow(sheet, wb, lastRow + 3, "⚠️ Errors", error, grayStyle);

            try (FileOutputStream fos = new FileOutputStream(outputPath)) {
                wb.write(fos);
            }
        }
    }

    private static CellStyle colorStyle(XSSFWorkbook wb, IndexedColors color) {
        CellStyle style = wb.createCellStyle();
        style.setFillForegroundColor(color.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        Font font = wb.createFont();
        font.setFontName("Arial");
        font.setFontHeightInPoints((short) 10);
        style.setFont(font);
        return style;
    }

    private static void addSummaryRow(Sheet sheet, XSSFWorkbook wb,
                                      int rowIdx, String label, long count, CellStyle style) {
        Row row = sheet.createRow(rowIdx);
        Font boldFont = wb.createFont();
        boldFont.setBold(true);
        boldFont.setFontName("Arial");
        boldFont.setFontHeightInPoints((short) 10);
        CellStyle boldStyle = wb.createCellStyle();
        boldStyle.cloneStyleFrom(style);
        boldStyle.setFont(boldFont);

        Cell c1 = row.createCell(0); c1.setCellValue(label); c1.setCellStyle(boldStyle);
        Cell c2 = row.createCell(1); c2.setCellValue(count); c2.setCellStyle(boldStyle);
    }
}