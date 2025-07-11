package com.example.managementSystem.controller;

import com.example.managementSystem.model.MunicipalRecord;
import com.example.managementSystem.repository.MunicipalRecordRepository;

import jakarta.servlet.http.HttpServletResponse;

import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.function.Function;
import java.util.stream.Collectors;

import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
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
@CrossOrigin(origins = "https://karyaprabandhan.vercel.app")
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
        public ResponseEntity<Map<String, Object>> listRecords(
                        @RequestParam(required = false) String wardNo,
                        @RequestParam(required = false) String status,
                        @RequestParam(required = false) String dateFrom,
                        @RequestParam(required = false) String dateTo,
                        @RequestParam(required = false) String madName,
                        @RequestParam(required = false) String vidhansabhaName,
                        @RequestParam(required = false) String workName,
                        @RequestParam(required = false) String approvedAmount,
                        @RequestParam(defaultValue = "0") int page, // page number
                        @RequestParam(defaultValue = "5") int size // page size
        ) {
                List<MunicipalRecord> records = municipalRecordRepository.findAll();

                List<MunicipalRecord> filtered = records.stream()
                                .filter(r -> wardNo == null || r.getWardNo().equals(wardNo))
                                .filter(r -> status == null || (r.getPhysicalStatus() != null
                                                && r.getPhysicalStatus().equals(status)))
                                .filter(r -> madName == null
                                                || (r.getMadName() != null && r.getMadName().equals(madName)))
                                .filter(r -> vidhansabhaName == null
                                                || (r.getVidhansabhaName() != null
                                                                && r.getVidhansabhaName().equals(vidhansabhaName)))
                                .filter(r -> workName == null
                                                || (r.getWorkName() != null && r.getWorkName().contains(workName)))
                                .filter(r -> approvedAmount == null
                                                || (r.getSanctionAmount() != null
                                                                && r.getSanctionAmount().contains(approvedAmount)))
                                .filter(r -> {
                                        if (dateFrom == null)
                                                return true;
                                        return r.getSanctionDate() != null
                                                        && r.getSanctionDate().compareTo(dateFrom) >= 0;
                                })
                                .filter(r -> {
                                        if (dateTo == null)
                                                return true;
                                        return r.getSanctionDate() != null
                                                        && r.getSanctionDate().compareTo(dateTo) <= 0;
                                })
                                .toList();

                // Pagination logic
                int start = Math.min(page * size, filtered.size());
                int end = Math.min(start + size, filtered.size());
                List<MunicipalRecord> paginated = filtered.subList(start, end);

                Map<String, Object> response = new HashMap<>();
                response.put("records", paginated);
                response.put("currentPage", page);
                response.put("pageSize", size);
                response.put("totalRecords", filtered.size());
                response.put("totalPages", (int) Math.ceil((double) filtered.size() / size));

                return ResponseEntity.ok(response);
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

        // @GetMapping("/export")
        // public void exportFilteredExcel(
        // @RequestParam(required = false) String wardNo,
        // @RequestParam(required = false) String status,
        // @RequestParam(required = false) String dateFrom,
        // @RequestParam(required = false) String dateTo,
        // @RequestParam(required = false) String madName,
        // @RequestParam(required = false) String vidhansabhaName,
        // @RequestParam(required = false) String workName,
        // @RequestParam(required = false) String approvedAmount,
        // @RequestParam(defaultValue = "0") int page,
        // @RequestParam(defaultValue = "5") int size,
        // HttpServletResponse response) throws IOException {
        // List<MunicipalRecord> allRecords = municipalRecordRepository.findAll();

        // // Filtering
        // List<MunicipalRecord> filtered = allRecords.stream()
        // .filter(r -> wardNo == null || r.getWardNo().equals(wardNo))
        // .filter(r -> status == null || (r.getPhysicalStatus() != null &&
        // r.getPhysicalStatus().equals(status)))
        // .filter(r -> madName == null || (r.getMadName() != null &&
        // r.getMadName().equals(madName)))
        // .filter(r -> vidhansabhaName == null
        // || (r.getVidhansabhaName() != null &&
        // r.getVidhansabhaName().equals(vidhansabhaName)))
        // .filter(r -> workName == null || (r.getWorkName() != null &&
        // r.getWorkName().contains(workName)))
        // .filter(r -> approvedAmount == null
        // || (r.getSanctionAmount() != null &&
        // r.getSanctionAmount().contains(approvedAmount)))
        // .filter(r -> dateFrom == null
        // || (r.getSanctionDate() != null
        // && r.getSanctionDate().compareTo(dateFrom) >= 0))
        // .filter(r -> dateTo == null
        // || (r.getSanctionDate() != null
        // && r.getSanctionDate().compareTo(dateTo) <= 0))
        // .toList();

        // // Pagination
        // int fromIndex = Math.min(page * size, filtered.size());
        // int toIndex = Math.min(fromIndex + size, filtered.size());
        // List<MunicipalRecord> paginated = filtered.subList(fromIndex, toIndex);

        // // Export paginated filtered data
        // exportToExcelInternal(response, paginated);
        // }

        // private void exportToExcelInternal(HttpServletResponse response,
        // List<MunicipalRecord> records) throws IOException {
        // // List<MunicipalRecord> records = municipalRecordRepository.findAll();
        // XSSFWorkbook workbook = new XSSFWorkbook();
        // XSSFSheet sheet = workbook.createSheet("Records");

        // // Fonts
        // Font boldFont = workbook.createFont();
        // boldFont.setFontName("Kruti Dev 010");
        // boldFont.setFontHeightInPoints((short) 12);
        // boldFont.setBold(true);

        // Font normalFont = workbook.createFont();
        // normalFont.setFontName("Kruti Dev 010");
        // normalFont.setFontHeightInPoints((short) 11);

        // // Styles without border (for top heading)
        // CellStyle headingStyle = workbook.createCellStyle();
        // headingStyle.setFont(boldFont);
        // headingStyle.setAlignment(HorizontalAlignment.CENTER);
        // headingStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        // headingStyle.setWrapText(true);

        // // Styles with border (for table)
        // CellStyle boldCenter = workbook.createCellStyle();
        // boldCenter.setFont(boldFont);
        // boldCenter.setAlignment(HorizontalAlignment.CENTER);
        // boldCenter.setVerticalAlignment(VerticalAlignment.CENTER);
        // boldCenter.setWrapText(true);
        // applyAllBorders(boldCenter);

        // CellStyle normalLeft = workbook.createCellStyle();
        // normalLeft.setFont(normalFont);
        // normalLeft.setAlignment(HorizontalAlignment.LEFT);
        // normalLeft.setVerticalAlignment(VerticalAlignment.TOP);
        // normalLeft.setWrapText(true);
        // applyAllBorders(normalLeft);

        // // Top Heading Block (no borders)
        // int rowIdx = 0;
        // String[] topHeaders = {
        // "dk;kZy; uxj ikfyd fuxe] jk;iqj ",
        // "tksu Øekad&03",
        // "fofHkUu enksa esa Lohd`r izxfrjr ,oa vizkjaHk dk;ksZ dh tkudkjh ",
        // // "jkuh ykSfdy cky okMZ dz- 10",
        // // "fnukad 01-04-2025 dh fLFkfr esa"
        // };
        // for (String heading : topHeaders) {
        // Row row = sheet.createRow(rowIdx++);
        // Cell cell = row.createCell(0);
        // cell.setCellValue(heading);
        // cell.setCellStyle(headingStyle);
        // sheet.addMergedRegion(new CellRangeAddress(row.getRowNum(), row.getRowNum(),
        // 0, 33));
        // }

        // // Header rows with borders
        // Row header1 = sheet.createRow(rowIdx++);
        // Row header2 = sheet.createRow(rowIdx++);

        // // Multi-Row Header
        // createMergedCell(sheet, header1, 0, 0, 2, "l-Ø-", boldCenter);
        // createMergedCell(sheet, header1, 1, 1, 2, "foRrh; o\"kZ", boldCenter);
        // createMergedCell(sheet, header1, 2, 2, 2, "fo/kkulHkk {ks= dk uke",
        // boldCenter);
        // createMergedCell(sheet, header1, 3, 3, 2, "en dk uke", boldCenter);
        // createMergedCell(sheet, header1, 4, 4, 2, "okMZ dk uke ,oa Ø-", boldCenter);
        // createMergedCell(sheet, header1, 5, 5, 2, "dk;Z dk uke", boldCenter);
        // createMergedCell(sheet, header1, 6, 6, 2, "Lohd`fr vkns'k Ø- ,oa fnukad",
        // boldCenter);

        // createMergedCell(sheet, header1, 7, 8, 1, "jkf'k", boldCenter);
        // createMergedCell(sheet, header2, 7, 7, 1, "iznk; jkf'k", boldCenter);
        // createMergedCell(sheet, header2, 8, 8, 1, "Lohd`r jkf'k", boldCenter);

        // createMergedCell(sheet, header1, 9, 11, 1, "fufonk", boldCenter);
        // createMergedCell(sheet, header2, 9, 9, 1, "izFke fufonk Ø- ,oa fnukad",
        // boldCenter);
        // createMergedCell(sheet, header2, 10, 10, 1, "f}rh; fufonk Ø- ,oa fnukad",
        // boldCenter);
        // createMergedCell(sheet, header2, 11, 11, 1, "r`rh; fufonk Ø- ,oa fnukad",
        // boldCenter);

        // createMergedCell(sheet, header1, 12, 12, 2, "fufonk nj", boldCenter);
        // createMergedCell(sheet, header1, 13, 13, 2, "fufonk i'pkr~ Lohd`r jkf'k",
        // boldCenter);
        // createMergedCell(sheet, header1, 14, 14, 2, "Bsdsnkj dk uke", boldCenter);
        // createMergedCell(sheet, header1, 15, 15, 2, "lEidZ fooj.k", boldCenter);
        // createMergedCell(sheet, header1, 16, 16, 2, "vuqca/k Ø- ,oa fnukad",
        // boldCenter);
        // createMergedCell(sheet, header1, 17, 17, 2, "dk;Z izkjaHk dk okLrfod fnukad",
        // boldCenter);
        // createMergedCell(sheet, header1, 18, 18, 2, "dk;Z lekfIr dk okLrfod fnukad",
        // boldCenter);

        // createMergedCell(sheet, header1, 19, 23, 1, "dqy O;; jkf'k", boldCenter);
        // createMergedCell(sheet, header2, 19, 19, 1, "izFke py ns;d", boldCenter);
        // createMergedCell(sheet, header2, 20, 20, 1, "f}rh; py ns;d", boldCenter);
        // createMergedCell(sheet, header2, 21, 21, 1, "r`rh; ,oa vafre", boldCenter);
        // createMergedCell(sheet, header2, 22, 22, 1, ";ksx", boldCenter);
        // createMergedCell(sheet, header2, 23, 23, 1, "'ks\"k jkf'k", boldCenter);

        // createMergedCell(sheet, header1, 24, 26, 1, "dk;Z dh HkkSfrd fLFkfr",
        // boldCenter);
        // createMergedCell(sheet, header2, 24, 24, 1, "vizkjaHk izxfr iw.kZ",
        // boldCenter);
        // createMergedCell(sheet, header2, 25, 25, 1, "izkDdyu vuqlkj yEckbZ@{ks=Qy",
        // boldCenter);
        // createMergedCell(sheet, header2, 26, 26, 1, "dqy orZeku okLrfod
        // HkkSfrdmiyfC/k", boldCenter);

        // createMergedCell(sheet, header1, 27, 27, 2, "fjekdZ", boldCenter);

        // // Data rows
        // int serial = 1;
        // for (MunicipalRecord r : records) {
        // Row row = sheet.createRow(rowIdx++);
        // int col = 0;
        // row.createCell(col++).setCellValue(serial++);
        // row.createCell(col++).setCellValue(r.getFinancialYear());
        // row.createCell(col++).setCellValue(r.getVidhansabhaName());
        // row.createCell(col++).setCellValue(r.getMadName());
        // row.createCell(col++).setCellValue(r.getWardNo());
        // row.createCell(col++).setCellValue(r.getWorkName());
        // row.createCell(col++).setCellValue(r.getSanctionedNumber());
        // row.createCell(col++).setCellValue(r.getAllotedAmount());
        // row.createCell(col++).setCellValue(r.getSanctionAmount());
        // row.createCell(col++).setCellValue(r.getTender1Number() + " " +
        // r.getTender1Date());
        // row.createCell(col++).setCellValue(r.getTender2Number() + " " +
        // r.getTender2Date());
        // row.createCell(col++).setCellValue(r.getTender3Number() + " " +
        // r.getTender3Date());
        // row.createCell(col++).setCellValue(r.getTenderRate());
        // row.createCell(col++).setCellValue(r.getPostTenderApprovedAmount());
        // row.createCell(col++).setCellValue(r.getContractorName());
        // row.createCell(col++).setCellValue(r.getContractorMobile());
        // row.createCell(col++).setCellValue(r.getAgreementNumber() + " " +
        // r.getAgreementDate());
        // row.createCell(col++).setCellValue(r.getActualStartDate());
        // row.createCell(col++).setCellValue(r.getActualCompletionDate());
        // row.createCell(col++).setCellValue(r.getFirstPayment());
        // row.createCell(col++).setCellValue(r.getSecondPayment());
        // row.createCell(col++).setCellValue(r.getFinalPayment());
        // row.createCell(col++).setCellValue(r.getTotalPayment());
        // row.createCell(col++).setCellValue(r.getRemainingAmount());
        // row.createCell(col++).setCellValue(r.getPhysicalStatus());
        // row.createCell(col++).setCellValue(r.getEstimatedPhysicalAchievement());
        // row.createCell(col++).setCellValue(r.getActualPhysicalAchievement());
        // row.createCell(col++).setCellValue(r.getRemarks());

        // for (int i = 0; i < col; i++) {
        // row.getCell(i).setCellStyle(normalLeft);
        // }
        // }

        // // Set column widths
        // int[] columnWidths = {
        // 900, 2000, 4000, 2500, 1500, 7000, 3000,
        // 2000, 2000, 3000, 3000, 3000,
        // 1000, 2000, 2000, 2000, 4000,
        // 1500, 1500,
        // 2000, 2000, 2000, 2000, 2000,
        // 2000, 2000, 2000,
        // 3000
        // };
        // for (int i = 0; i < columnWidths.length; i++) {
        // sheet.setColumnWidth(i, columnWidths[i]);
        // }

        // workbook.write(response.getOutputStream());
        // workbook.close();
        // }

        // // Add border to all 4 sides of a cell
        // private void applyAllBorders(CellStyle style) {
        // style.setBorderTop(BorderStyle.THIN);
        // style.setBorderBottom(BorderStyle.THIN);
        // style.setBorderLeft(BorderStyle.THIN);
        // style.setBorderRight(BorderStyle.THIN);
        // }

        // // Merge cells + apply border to merged area
        // private void createMergedCell(XSSFSheet sheet, Row row, int colStart, int
        // colEnd, int rowSpan, String text,
        // CellStyle style) {
        // Cell cell = row.createCell(colStart);
        // cell.setCellValue(text);
        // cell.setCellStyle(style);

        // int rowStart = row.getRowNum();
        // int rowEnd = rowStart + rowSpan - 1;

        // if (colStart != colEnd || rowSpan > 1) {
        // sheet.addMergedRegion(new CellRangeAddress(rowStart, rowEnd, colStart,
        // colEnd));
        // for (int r = rowStart; r <= rowEnd; r++) {
        // Row currentRow = sheet.getRow(r);
        // if (currentRow == null)
        // currentRow = sheet.createRow(r);
        // for (int c = colStart; c <= colEnd; c++) {
        // Cell currentCell = currentRow.getCell(c);
        // if (currentCell == null)
        // currentCell = currentRow.createCell(c);
        // currentCell.setCellStyle(style);
        // }
        // }
        // }

        // }

       
        // ✅ Updated Excel Export API - Only Exports Selected Columns with Full Header
        // Structure

        @GetMapping("/export")
        public void exportFilteredExcel(
                        @RequestParam(required = false) String wardNo,
                        @RequestParam(required = false) String status,
                        @RequestParam(required = false) String dateFrom,
                        @RequestParam(required = false) String dateTo,
                        @RequestParam(required = false) String madName,
                        @RequestParam(required = false) String vidhansabhaName,
                        @RequestParam(required = false) String workName,
                        @RequestParam(required = false) String approvedAmount,
                        @RequestParam(defaultValue = "0") int page,
                        @RequestParam(defaultValue = "5") int size,
                        @RequestParam(required = false) List<String> columns,
                        HttpServletResponse response) throws IOException {

                List<MunicipalRecord> allRecords = municipalRecordRepository.findAll();

                List<MunicipalRecord> filtered = allRecords.stream()
                                .filter(r -> wardNo == null || r.getWardNo().equals(wardNo))
                                .filter(r -> status == null || (r.getPhysicalStatus() != null
                                                && r.getPhysicalStatus().equals(status)))
                                .filter(r -> madName == null
                                                || (r.getMadName() != null && r.getMadName().equals(madName)))
                                .filter(r -> vidhansabhaName == null || (r.getVidhansabhaName() != null
                                                && r.getVidhansabhaName().equals(vidhansabhaName)))
                                .filter(r -> workName == null
                                                || (r.getWorkName() != null && r.getWorkName().contains(workName)))
                                .filter(r -> approvedAmount == null || (r.getSanctionAmount() != null
                                                && r.getSanctionAmount().contains(approvedAmount)))
                                .filter(r -> dateFrom == null || (r.getSanctionDate() != null
                                                && r.getSanctionDate().compareTo(dateFrom) >= 0))
                                .filter(r -> dateTo == null || (r.getSanctionDate() != null
                                                && r.getSanctionDate().compareTo(dateTo) <= 0))
                                .collect(Collectors.toList());

                int fromIndex = Math.min(page * size, filtered.size());
                int toIndex = Math.min(fromIndex + size, filtered.size());
                List<MunicipalRecord> paginated = filtered.subList(fromIndex, toIndex);

                exportSelectedFieldsToExcel(response, paginated, columns);
        }

        private void exportSelectedFieldsToExcel(HttpServletResponse response, List<MunicipalRecord> records,
                        List<String> selectedColumns) throws IOException {

                XSSFWorkbook workbook = new XSSFWorkbook();
                XSSFSheet sheet = workbook.createSheet("Records");

                // Fonts
                Font krutiFont = workbook.createFont();
                krutiFont.setFontName("Kruti Dev 010");
                krutiFont.setFontHeightInPoints((short) 11);

                Font krutiBoldFont = workbook.createFont();
                krutiBoldFont.setFontName("Kruti Dev 010");
                krutiBoldFont.setFontHeightInPoints((short) 11);
                krutiBoldFont.setBold(true);

                Font krutiTitleFont = workbook.createFont();
                krutiTitleFont.setFontName("Kruti Dev 010");
                krutiTitleFont.setFontHeightInPoints((short) 14);
                krutiTitleFont.setBold(true);

                Font normalFont = workbook.createFont();
                normalFont.setFontName("Arial");
                normalFont.setFontHeightInPoints((short) 11);

                CellStyle normalStyle = workbook.createCellStyle();
                normalStyle.setFont(krutiFont);
                normalStyle.setWrapText(true);
                normalStyle.setAlignment(HorizontalAlignment.LEFT);
                normalStyle.setVerticalAlignment(VerticalAlignment.TOP);
                applyAllBorders(normalStyle);

                CellStyle headerStyle = workbook.createCellStyle();
                headerStyle.cloneStyleFrom(normalStyle);
                headerStyle.setFont(krutiBoldFont);
                headerStyle.setAlignment(HorizontalAlignment.CENTER);
                headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
                headerStyle.setWrapText(true);
                applyAllBorders(headerStyle);

                CellStyle titleStyle = workbook.createCellStyle();
                titleStyle.cloneStyleFrom(headerStyle);
                titleStyle.setFont(krutiTitleFont);

                CellStyle amountStyle = workbook.createCellStyle();
                amountStyle.setFont(normalFont);
                amountStyle.setWrapText(true);
                amountStyle.setAlignment(HorizontalAlignment.RIGHT);
                amountStyle.setVerticalAlignment(VerticalAlignment.TOP);
                applyAllBorders(amountStyle);

                CellStyle titleNoBorderStyle = workbook.createCellStyle();
                titleNoBorderStyle.cloneStyleFrom(titleStyle);
                titleNoBorderStyle.setBorderTop(BorderStyle.NONE);
                titleNoBorderStyle.setBorderBottom(BorderStyle.NONE);
                titleNoBorderStyle.setBorderLeft(BorderStyle.NONE);
                titleNoBorderStyle.setBorderRight(BorderStyle.NONE);

                // Subtitle row without border
                CellStyle subtitleNoBorderStyle = workbook.createCellStyle();
                subtitleNoBorderStyle.cloneStyleFrom(headerStyle);
                subtitleNoBorderStyle.setBorderTop(BorderStyle.NONE);
                subtitleNoBorderStyle.setBorderBottom(BorderStyle.NONE);
                subtitleNoBorderStyle.setBorderLeft(BorderStyle.NONE);
                subtitleNoBorderStyle.setBorderRight(BorderStyle.NONE);

                record ColumnMeta(String field, String krutiLabel, String parent) {
                }

                List<ColumnMeta> allColumns = List.of(
                                new ColumnMeta("serial", "l-Ø-", null),
                                new ColumnMeta("financialYear", "foRrh; o\"kZ", null),
                                new ColumnMeta("vidhansabhaName", "fo/kkulHkk {ks= dk uke", null),
                                new ColumnMeta("madName", "en dk uke", null),
                                new ColumnMeta("wardNo", "okMZ dk uke ,oa Ø-", null),
                                new ColumnMeta("workName", "dk;Z dk uke", null),
                                new ColumnMeta("sanctionedNumber", "Lohd`fr vkns'k Ø-", null),

                                new ColumnMeta("allotedAmount", "iznk; jkf'k", "jkf'k"),
                                new ColumnMeta("sanctionAmount", "Lohd`r jkf'k", "jkf'k"),

                                new ColumnMeta("tender1Combined", "izFke fufonk Ø- ,oa fnukad", "fufonk"),
                                new ColumnMeta("tender2Combined", "f}rh; fufonk Ø- ,oa fnukad", "fufonk"),
                                new ColumnMeta("tender3Combined", "r`rh; fufonk Ø- ,oa fnukad", "fufonk"),
                                new ColumnMeta("tenderRate", "fufonk nj", "fufonk"),

                                new ColumnMeta("postTenderApprovedAmount", "fufonk i'pkr~ Lohd`r jkf'k", null),
                                new ColumnMeta("contractorName", "Bsdsnkj dk uke", null),
                                new ColumnMeta("contractorMobile", "lEidZ fooj.k", null),

                                new ColumnMeta("agreementCombined", "vuqca/k Ø- ,oa fnukad", null),
                                new ColumnMeta("workOrderCombined", "dk;Z vkns'k Ø- ,oa fnukad", null),
                                new ColumnMeta("sectionCombined", "Lohd`fr vkns'k Ø- ,oa fnukad", null),

                                new ColumnMeta("actualStartDate", "dk;Z 'kq: dh fnukad", null),
                                new ColumnMeta("actualCompletionDate", "dk;Z lekfIr dh fnukad", null),

                                new ColumnMeta("firstPayment", "izFke py ns;d", "dqy O;; jkf'k"),
                                new ColumnMeta("secondPayment", "f}rh; py ns;d", "dqy O;; jkf'k"),
                                new ColumnMeta("finalPayment", "vafre py ns;d", "dqy O;; jkf'k"),
                                new ColumnMeta("totalPayment", ";ksx jkf'k", "dqy O;; jkf'k"),
                                new ColumnMeta("remainingAmount", "'ks\"k jkf'k", "dqy O;; jkf'k"),

                                new ColumnMeta("physicalProgressStatus", "vizkjaHk izxfr iw.kZ",
                                                "dk;Z dh HkkSfrd fLFkfr"),
                                new ColumnMeta("estimatedPhysicalAchievement", "izkDdyu vuqlkj yEckbZ@{ks=Qy",
                                                "dk;Z dh HkkSfrd fLFkfr"),
                                new ColumnMeta("actualPhysicalAchievement", "dqy orZeku okLrfod HkkSfrdmiyfC/k",
                                                "dk;Z dh HkkSfrd fLFkfr"),

                                new ColumnMeta("remarks", "fjekdZ", null));

                Set<String> selectedSet = new HashSet<>(selectedColumns != null ? selectedColumns : List.of());
                selectedSet.add("serial");

                if (selectedSet.contains("tender1Number") || selectedSet.contains("tender1Date"))
                        selectedSet.add("tender1Combined");
                if (selectedSet.contains("tender2Number") || selectedSet.contains("tender2Date"))
                        selectedSet.add("tender2Combined");
                if (selectedSet.contains("tender3Number") || selectedSet.contains("tender3Date"))
                        selectedSet.add("tender3Combined");
                if (selectedSet.contains("agreementNumber") || selectedSet.contains("agreementDate")) {
                        selectedSet.remove("agreementNumber");
                        selectedSet.remove("agreementDate");
                        selectedSet.add("agreementCombined");
                }
                if (selectedSet.contains("workOrderNumber") || selectedSet.contains("workOrderDate")) {
                        selectedSet.remove("workOrderNumber");
                        selectedSet.remove("workOrderDate");
                        selectedSet.add("workOrderCombined");
                }
                if (selectedSet.contains("sanctionedNumber") || selectedSet.contains("sanctionDate")) {
                        selectedSet.remove("sanctionedNumber");
                        selectedSet.remove("sanctionDate");
                        selectedSet.add("sectionCombined");
                }
                if (selectedSet.contains("physicalStatus") || selectedSet.contains("progressPercentage")) {
                        selectedSet.remove("physicalStatus");
                        selectedSet.remove("progressPercentage");
                        selectedSet.add("physicalProgressStatus");
                }

                List<ColumnMeta> selectedMetas = allColumns.stream()
                                .filter(meta -> selectedSet.contains(meta.field()))
                                .toList();

                Map<String, List<ColumnMeta>> grouped = selectedMetas.stream()
                                .collect(Collectors.groupingBy(
                                                meta -> meta.parent() == null ? meta.field() : meta.parent(),
                                                LinkedHashMap::new,
                                                Collectors.toList()));

                int rowIdx = 0;

                int totalCols = grouped.values().stream().mapToInt(List::size).sum();

                // First title row
                // First title row
                Row titleRow1 = sheet.createRow(rowIdx++);
                Cell titleCell1 = titleRow1.createCell(0);
                titleCell1.setCellValue("dk;kZy; uxj ikfyd fuxe] jk;iqj");
                titleCell1.setCellStyle(titleNoBorderStyle);
                sheet.addMergedRegion(new CellRangeAddress(rowIdx - 1, rowIdx - 1, 0, totalCols - 1));

                // Second title row
                Row titleRow2 = sheet.createRow(rowIdx++);
                Cell titleCell2 = titleRow2.createCell(0);
                titleCell2.setCellValue("tksu Øekad&03");
                titleCell2.setCellStyle(titleNoBorderStyle);
                sheet.addMergedRegion(new CellRangeAddress(rowIdx - 1, rowIdx - 1, 0, totalCols - 1));

                // Third title row (subtitle)
                Row titleRow3 = sheet.createRow(rowIdx++);
                Cell titleCell3 = titleRow3.createCell(0);
                titleCell3.setCellValue("fofHkUu enksa esa Lohd`r izxfrjr ,oa vizkjaHk dk;ksZ dh tkudkjh");
                titleCell3.setCellStyle(subtitleNoBorderStyle);
                sheet.addMergedRegion(new CellRangeAddress(rowIdx - 1, rowIdx - 1, 0, totalCols - 1));

                Row header1 = sheet.createRow(rowIdx++);
                Row header2 = sheet.createRow(rowIdx++);
                int colIdx = 0;

                for (Map.Entry<String, List<ColumnMeta>> entry : grouped.entrySet()) {
                        String group = entry.getKey();
                        List<ColumnMeta> fields = entry.getValue();

                        if (fields.size() == 1 && fields.get(0).parent() == null) {
                                createMergedCell(sheet, header1, colIdx, colIdx, 2, fields.get(0).krutiLabel(),
                                                headerStyle);
                                colIdx++;
                        } else {
                                createMergedCell(sheet, header1, colIdx, colIdx + fields.size() - 1, 1, group,
                                                headerStyle);
                                for (ColumnMeta meta : fields) {
                                        Cell cell = header2.createCell(colIdx++);
                                        cell.setCellValue(meta.krutiLabel());
                                        cell.setCellStyle(headerStyle);
                                }
                        }
                }

                int serial = 1;
                for (MunicipalRecord record : records) {
                        Row row = sheet.createRow(rowIdx++);
                        int dataCol = 0;

                        for (Map.Entry<String, List<ColumnMeta>> entry : grouped.entrySet()) {
                                for (ColumnMeta meta : entry.getValue()) {
                                        Cell cell = row.createCell(dataCol++);
                                        Object value = switch (meta.field()) {
                                                case "serial" -> serial++;
                                                case "tender1Combined" -> combineFields(record.getTender1Number(),
                                                                record.getTender1Date());
                                                case "tender2Combined" -> combineFields(record.getTender2Number(),
                                                                record.getTender2Date());
                                                case "tender3Combined" -> combineFields(record.getTender3Number(),
                                                                record.getTender3Date());
                                                case "agreementCombined" -> combineFields(record.getAgreementNumber(),
                                                                record.getAgreementDate());
                                                case "workOrderCombined" -> combineFields(record.getWorkOrderNumber(),
                                                                record.getWorkOrderDate());
                                                case "sectionCombined" -> combineFields(record.getSanctionedNumber(),
                                                                record.getSanctionDate());
                                                case "physicalProgressStatus" -> {
                                                        String status = record.getPhysicalStatus();
                                                        String progress = record.getProgressPercentage();
                                                        yield (status != null ? status : "")
                                                                        + (progress != null ? " (" + progress + "%)"
                                                                                        : "");
                                                }
                                                case "tenderRate" -> record.getTenderRate();
                                                default -> getFieldValue(record, meta.field());
                                        };
                                        cell.setCellValue(value != null ? value.toString() : "");
                                        cell.setCellStyle(isAmountField(meta.field()) ? amountStyle : normalStyle);
                                }
                        }
                }

                // for (int i = 0; i < colIdx; i++) {
                // sheet.autoSizeColumn(i);
                // }

                // After writing headers and before writing data or at the end before response
                Map<String, Integer> columnWidths = Map.ofEntries(
                                Map.entry("serial", 900),
                                Map.entry("financialYear", 2000),
                                Map.entry("vidhansabhaName", 4000),
                                Map.entry("madName", 2500),
                                Map.entry("wardNo", 1500),
                                Map.entry("workName", 7000),
                                Map.entry("sanctionedNumber", 3000),
                                Map.entry("allotedAmount", 2000),
                                Map.entry("sanctionAmount", 2000),
                                Map.entry("tender1Combined", 3000),
                                Map.entry("tender2Combined", 3000),
                                Map.entry("tender3Combined", 3000),
                                Map.entry("tenderRate", 1000),
                                Map.entry("contractorName", 2000),
                                Map.entry("contractorMobile", 2000),
                                Map.entry("agreementCombined", 2000),
                                Map.entry("workOrderCombined", 4000),
                                Map.entry("sectionCombined", 1500),
                                Map.entry("actualStartDate", 1500),
                                Map.entry("actualCompletionDate", 2000),
                                Map.entry("firstPayment", 2000),
                                Map.entry("secondPayment", 2000),
                                Map.entry("finalPayment", 2000),
                                Map.entry("totalPayment", 2000),
                                Map.entry("remainingAmount", 2000),
                                Map.entry("physicalProgressStatus", 2000),
                                Map.entry("estimatedPhysicalAchievement", 2000),
                                Map.entry("actualPhysicalAchievement", 2000),
                                Map.entry("remarks", 3000));

                int columnIndex = 0;
                for (Map.Entry<String, List<ColumnMeta>> entry : grouped.entrySet()) {
                        for (ColumnMeta meta : entry.getValue()) {
                                int width = columnWidths.getOrDefault(meta.field(), 3000); // Default to 3000 if not
                                                                                           // listed
                                sheet.setColumnWidth(columnIndex++, width);
                        }
                }

                response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
                response.setHeader("Content-Disposition", "attachment; filename=\"नगर_निगम_रिकॉर्ड.xlsx\"");
                workbook.write(response.getOutputStream());
                workbook.close();
        }

        private void applyAllBorders(CellStyle style) {
                style.setBorderTop(BorderStyle.THIN);
                style.setBorderBottom(BorderStyle.THIN);
                style.setBorderLeft(BorderStyle.THIN);
                style.setBorderRight(BorderStyle.THIN);
        }

        private String combineFields(String number, String date) {
                if (number == null && date == null)
                        return "";
                if (number == null)
                        return date;
                if (date == null)
                        return number;
                return number + " / " + date;
        }

        private Object getFieldValue(MunicipalRecord record, String field) {
                try {
                        Field f = MunicipalRecord.class.getDeclaredField(field);
                        f.setAccessible(true);
                        return f.get(record);
                } catch (Exception e) {
                        return "";
                }
        }

        private boolean isAmountField(String field) {
                return List.of(
                                "sanctionAmount",
                                "allotedAmount",
                                "postTenderApprovedAmount",
                                "firstPayment",
                                "secondPayment",
                                "finalPayment",
                                "totalPayment",
                                "remainingAmount",
                                "tenderRate").contains(field);
        }

        private void createMergedCell(XSSFSheet sheet, Row row, int colStart, int colEnd, int rowSpan,
                        String text, CellStyle style) {
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
