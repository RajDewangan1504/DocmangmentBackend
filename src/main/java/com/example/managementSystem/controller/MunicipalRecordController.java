package com.example.managementSystem.controller;

import com.example.managementSystem.model.MunicipalRecord;
import com.example.managementSystem.repository.MunicipalRecordRepository;

import jakarta.servlet.http.HttpServletResponse;

import java.io.IOException;
import java.util.List;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;
import org.springframework.http.ResponseEntity;
import org.springframework.http.HttpStatus;

@RestController
@RequestMapping("/api/municipal-records")
// @CrossOrigin(origins = "*")
@CrossOrigin(origins = "http://localhost:8081")
public class MunicipalRecordController {

    @Autowired
    private MunicipalRecordRepository municipalRecordRepository;

    @PostMapping("/create")
    public ResponseEntity<MunicipalRecord> createRecord(@RequestBody MunicipalRecord record) {
        // System.out.println("➡️ Incoming data: " + record); // DEBUG PRINT
        MunicipalRecord saved = municipalRecordRepository.save(record);
        return ResponseEntity.ok(saved);
    }

    @GetMapping("/list")
    public ResponseEntity<List<MunicipalRecord>> listRecords() {
        return ResponseEntity.ok(municipalRecordRepository.findAll());
    }

    @GetMapping("/{id}")
    public ResponseEntity<MunicipalRecord> getRecordById(@PathVariable String id) {
        return municipalRecordRepository.findById(id)
                .map(record -> ResponseEntity.ok(record))
                .orElse(ResponseEntity.status(HttpStatus.NOT_FOUND).body(null));
    }

    @PutMapping("/update/{id}")
    public ResponseEntity<MunicipalRecord> updateRecord(@PathVariable String id,
            @RequestBody MunicipalRecord updatedRecord) {
        return municipalRecordRepository.findById(id)
                .map(existing -> {
                    updatedRecord.setId(id); // Ensure ID remains unchanged
                    MunicipalRecord saved = municipalRecordRepository.save(updatedRecord);
                    return ResponseEntity.ok(saved);
                })
                .orElse(ResponseEntity.status(HttpStatus.NOT_FOUND).body(null));
    }

    @DeleteMapping("/delete/{id}")
    public ResponseEntity<String> deleteRecord(@PathVariable String id) {
        return municipalRecordRepository.findById(id)
                .map(record -> {
                    municipalRecordRepository.deleteById(id);
                    return ResponseEntity.ok("Record with ID " + id + " deleted successfully.");
                })
                .orElse(ResponseEntity.status(HttpStatus.NOT_FOUND)
                        .body("Record not found with ID " + id));
    }

    @GetMapping("/export")
    public void exportToExcel(HttpServletResponse response) throws IOException {
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        response.setHeader("Content-Disposition", "attachment; filename=municipal_records.xlsx");

        List<MunicipalRecord> records = municipalRecordRepository.findAll();
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Records");

        // Fonts
        Font boldFont = workbook.createFont();
        boldFont.setFontName("Kruti Dev 010");
        boldFont.setFontHeightInPoints((short) 12);
        boldFont.setBold(true);

        Font normalFont = workbook.createFont();
        normalFont.setFontName("Kruti Dev 010");
        normalFont.setFontHeightInPoints((short) 11);

        // Styles without border (for top heading)
        CellStyle headingStyle = workbook.createCellStyle();
        headingStyle.setFont(boldFont);
        headingStyle.setAlignment(HorizontalAlignment.CENTER);
        headingStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        headingStyle.setWrapText(true);

        // Styles with border (for table)
        CellStyle boldCenter = workbook.createCellStyle();
        boldCenter.setFont(boldFont);
        boldCenter.setAlignment(HorizontalAlignment.CENTER);
        boldCenter.setVerticalAlignment(VerticalAlignment.CENTER);
        boldCenter.setWrapText(true);
        applyAllBorders(boldCenter);

        CellStyle normalLeft = workbook.createCellStyle();
        normalLeft.setFont(normalFont);
        normalLeft.setAlignment(HorizontalAlignment.LEFT);
        normalLeft.setVerticalAlignment(VerticalAlignment.TOP);
        normalLeft.setWrapText(true);
        applyAllBorders(normalLeft);

        // Top Heading Block (no borders)
        int rowIdx = 0;
        String[] topHeaders = {
                "dk;kZy; uxj ikfyd fuxe] jk;iqj ",
                "tksu Øekad&03",
                "fofHkUu enksa esa Lohd`r izxfrjr ,oa vizkjaHk dk;ksZ dh tkudkjh  ",
                // "jkuh ykSfdy cky okMZ dz- 10",
                // "fnukad 01-04-2025 dh fLFkfr esa"
        };
        for (String heading : topHeaders) {
            Row row = sheet.createRow(rowIdx++);
            Cell cell = row.createCell(0);
            cell.setCellValue(heading);
            cell.setCellStyle(headingStyle);
            sheet.addMergedRegion(new CellRangeAddress(row.getRowNum(), row.getRowNum(), 0, 33));
        }

        // Header rows with borders
        Row header1 = sheet.createRow(rowIdx++);
        Row header2 = sheet.createRow(rowIdx++);

        // Multi-Row Header
        createMergedCell(sheet, header1, 0, 0, 2, "l-Ø-", boldCenter);
        createMergedCell(sheet, header1, 1, 1, 2, "foRrh; o\"kZ", boldCenter);
        createMergedCell(sheet, header1, 2, 2, 2, "fo/kkulHkk {ks= dk uke", boldCenter);
        createMergedCell(sheet, header1, 3, 3, 2, "en dk uke", boldCenter);
        createMergedCell(sheet, header1, 4, 4, 2, "okMZ dk uke ,oa Ø-", boldCenter);
        createMergedCell(sheet, header1, 5, 5, 2, "dk;Z dk uke", boldCenter);
        createMergedCell(sheet, header1, 6, 6, 2, "Lohd`fr vkns'k Ø- ,oa fnukad", boldCenter);

        createMergedCell(sheet, header1, 7, 8, 1, "jkf'k", boldCenter);
        createMergedCell(sheet, header2, 7, 7, 1, "iznk; jkf'k", boldCenter);
        createMergedCell(sheet, header2, 8, 8, 1, "Lohd`r jkf'k", boldCenter);

        createMergedCell(sheet, header1, 9, 11, 1, "fufonk", boldCenter);
        createMergedCell(sheet, header2, 9, 9, 1, "izFke fufonk Ø- ,oa fnukad", boldCenter);
        createMergedCell(sheet, header2, 10, 10, 1, "f}rh; fufonk Ø- ,oa fnukad", boldCenter);
        createMergedCell(sheet, header2, 11, 11, 1, "r`rh; fufonk Ø- ,oa fnukad", boldCenter);

        createMergedCell(sheet, header1, 12, 12, 2, "fufonk nj", boldCenter);
        createMergedCell(sheet, header1, 13, 13, 2, "fufonk i'pkr~ Lohd`r jkf'k", boldCenter);
        createMergedCell(sheet, header1, 14, 14, 2, "Bsdsnkj dk uke", boldCenter);
        createMergedCell(sheet, header1, 15, 15, 2, "lEidZ fooj.k", boldCenter);
        createMergedCell(sheet, header1, 16, 16, 2, "vuqca/k Ø- ,oa fnukad", boldCenter);
        createMergedCell(sheet, header1, 17, 17, 2, "dk;Z izkjaHk dk okLrfod fnukad", boldCenter);
        createMergedCell(sheet, header1, 18, 18, 2, "dk;Z lekfIr dk okLrfod fnukad", boldCenter);

        createMergedCell(sheet, header1, 19, 23, 1, "dqy O;; jkf'k", boldCenter);
        createMergedCell(sheet, header2, 19, 19, 1, "izFke py ns;d", boldCenter);
        createMergedCell(sheet, header2, 20, 20, 1, "f}rh; py ns;d", boldCenter);
        createMergedCell(sheet, header2, 21, 21, 1, "r`rh; ,oa vafre", boldCenter);
        createMergedCell(sheet, header2, 22, 22, 1, ";ksx", boldCenter);
        createMergedCell(sheet, header2, 23, 23, 1, "'ks\"k jkf'k", boldCenter);

        createMergedCell(sheet, header1, 24, 26, 1, "dk;Z dh HkkSfrd fLFkfr", boldCenter);
        createMergedCell(sheet, header2, 24, 24, 1, "vizkjaHk izxfr iw.kZ", boldCenter);
        createMergedCell(sheet, header2, 25, 25, 1, "izkDdyu vuqlkj yEckbZ@{ks=Qy", boldCenter);
        createMergedCell(sheet, header2, 26, 26, 1, "dqy orZeku okLrfod HkkSfrd miyfC/k", boldCenter);

        createMergedCell(sheet, header1, 27, 27, 2, "fjekdZ", boldCenter);

        // Data rows
        int serial = 1;
        for (MunicipalRecord r : records) {
            Row row = sheet.createRow(rowIdx++);
            int col = 0;
            row.createCell(col++).setCellValue(serial++);
            row.createCell(col++).setCellValue(r.getFinancialYear());
            row.createCell(col++).setCellValue(r.getVidhansabhaName());
            row.createCell(col++).setCellValue(r.getMadName());
            row.createCell(col++).setCellValue(r.getWardNo());
            row.createCell(col++).setCellValue(r.getWorkName());
            row.createCell(col++).setCellValue(r.getSanctionedNumber());
            row.createCell(col++).setCellValue(r.getAllotedAmount());
            row.createCell(col++).setCellValue(r.getSanctionAmount());
            row.createCell(col++).setCellValue(r.getTender1Number() + " " + r.getTender1Date());
            row.createCell(col++).setCellValue(r.getTender2Number() + " " + r.getTender2Date());
            row.createCell(col++).setCellValue(r.getTender3Number() + " " + r.getTender3Date());
            row.createCell(col++).setCellValue(r.getTenderRate());
            row.createCell(col++).setCellValue(r.getPostTenderApprovedAmount());
            row.createCell(col++).setCellValue(r.getContractorName());
            row.createCell(col++).setCellValue(r.getContractorMobile());
            row.createCell(col++).setCellValue(r.getAgreementNumber() + " " + r.getAgreementDate());
            row.createCell(col++).setCellValue(r.getActualStartDate());
            row.createCell(col++).setCellValue(r.getActualCompletionDate());
            row.createCell(col++).setCellValue(r.getFirstPayment());
            row.createCell(col++).setCellValue(r.getSecondPayment());
            row.createCell(col++).setCellValue(r.getFinalPayment());
            row.createCell(col++).setCellValue(r.getTotalPayment());
            row.createCell(col++).setCellValue(r.getRemainingAmount());
            row.createCell(col++).setCellValue(r.getPhysicalStatus());
            row.createCell(col++).setCellValue(r.getEstimatedPhysicalAchievement());
            row.createCell(col++).setCellValue(r.getActualPhysicalAchievement());
            row.createCell(col++).setCellValue(r.getRemarks());

            for (int i = 0; i < col; i++) {
                row.getCell(i).setCellStyle(normalLeft);
            }
        }

        // Set column widths
        int[] columnWidths = {
                1800, 3000, 5000, 4000, 4000, 7000, 5000,
                4000, 4000, 6000, 6000, 6000,
                3000, 4000, 4000, 4000, 6000,
                4000, 4000,
                4000, 4000, 4000, 4000, 4000,
                4000, 4000, 4000,
                6000
        };
        for (int i = 0; i < columnWidths.length; i++) {
            sheet.setColumnWidth(i, columnWidths[i]);
        }

        workbook.write(response.getOutputStream());
        workbook.close();
    }

    // Add border to all 4 sides of a cell
    private void applyAllBorders(CellStyle style) {
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
    }

    // Merge cells + apply border to merged area
    private void createMergedCell(XSSFSheet sheet, Row row, int colStart, int colEnd, int rowSpan, String text,
            CellStyle style) {
        Cell cell = row.createCell(colStart);
        cell.setCellValue(text);
        cell.setCellStyle(style);

        int rowStart = row.getRowNum();
        int rowEnd = rowStart + rowSpan - 1;

        if (colStart != colEnd || rowSpan > 1) {
            sheet.addMergedRegion(new CellRangeAddress(rowStart, rowEnd, colStart, colEnd));
            for (int r = rowStart; r <= rowEnd; r++) {
                Row currentRow = sheet.getRow(r);
                if (currentRow == null)
                    currentRow = sheet.createRow(r);
                for (int c = colStart; c <= colEnd; c++) {
                    Cell currentCell = currentRow.getCell(c);
                    if (currentCell == null)
                        currentCell = currentRow.createCell(c);
                    currentCell.setCellStyle(style);
                }
            }
        }
    }

}
