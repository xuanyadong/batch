package com.tjexcel.batch.service;

import com.tjexcel.batch.config.UpstreamDownstreamConfig;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.Instant;
import java.time.LocalDate;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.Date;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

/**
 * Generate upstream/downstream rows and append to a target workbook.
 */
@Service
public class UpstreamDownstreamService {

    private static final Logger log = LoggerFactory.getLogger(UpstreamDownstreamService.class);
    private static final DateTimeFormatter YEAR_MONTH_FMT = DateTimeFormatter.ofPattern("yyyy年MM月");

    private final UpstreamDownstreamConfig config;

    public UpstreamDownstreamService(UpstreamDownstreamConfig config) {
        this.config = config;
    }

    public int generate() throws IOException {
        Path dataFile = Paths.get(config.getDataPath()).toAbsolutePath();
        Path targetFile = Paths.get(config.getTargetFilePath()).toAbsolutePath();

        if (!Files.exists(dataFile)) {
            throw new FileNotFoundException("数据表不存在: " + dataFile);
        }

        Path parent = targetFile.getParent();
        if (parent != null) {
            Files.createDirectories(parent);
        }

        int sheetWritten = 0;
        int rowWritten = 0;

        try (InputStream dataIs = Files.newInputStream(dataFile);
             Workbook dataWb = openWorkbook(dataIs, dataFile);
             Workbook targetWb = openOrCreateWorkbook(targetFile)) {

            Sheet targetSheet = resolveTargetSheet(targetWb, config.getTargetSheetName());
            CellStyle headerStyle = createHeaderStyle(targetWb);
            CellStyle textStyle = createTextStyle(targetWb);

            ensureHeader(targetSheet, headerStyle);
            int rowCursor = findNextAppendRow(targetSheet);

            for (int si = 0; si < dataWb.getNumberOfSheets(); si++) {
                Sheet inSheet = dataWb.getSheetAt(si);
                if (inSheet == null) {
                    continue;
                }

                ParseResult parsed = parseSheet(inSheet);
                if (parsed.edges.isEmpty()) {
                    log.info("Sheet [{}] 没有有效供需关系，跳过", inSheet.getSheetName());
                    continue;
                }

                int before = rowCursor;
                rowCursor = appendParsedRows(targetSheet, rowCursor, parsed, textStyle);
                int appended = rowCursor - before;
                rowWritten += appended;
                sheetWritten++;

                log.info("Sheet [{}] 已追加 {} 行", inSheet.getSheetName(), appended);
            }

            for (int c = 0; c <= 4; c++) {
                targetSheet.autoSizeColumn(c);
                int width = targetSheet.getColumnWidth(c);
                targetSheet.setColumnWidth(c, Math.max(width, 18 * 256));
            }

            try (OutputStream os = Files.newOutputStream(targetFile)) {
                targetWb.write(os);
            }
        }

        log.info("上下游写入完成：追加 {} 个sheet，{} 行。目标文件: {}", sheetWritten, rowWritten, targetFile);
        return sheetWritten;
    }

    private Workbook openOrCreateWorkbook(Path targetFile) throws IOException {
        if (Files.exists(targetFile)) {
            try (InputStream is = Files.newInputStream(targetFile)) {
                return openWorkbook(is, targetFile);
            }
        }

        String lower = targetFile.getFileName().toString().toLowerCase();
        if (lower.endsWith(".xls")) {
            return new HSSFWorkbook();
        }
        return new XSSFWorkbook();
    }

    private Sheet resolveTargetSheet(Workbook wb, String configuredName) {
        String name = configuredName == null ? "" : configuredName.trim();
        if (!name.isEmpty()) {
            Sheet s = wb.getSheet(name);
            if (s != null) {
                return s;
            }
            return wb.createSheet(sanitizeSheetName(name));
        }

        if (wb.getNumberOfSheets() > 0) {
            return wb.getSheetAt(0);
        }
        return wb.createSheet("Sheet1");
    }

    private void ensureHeader(Sheet sheet, CellStyle headerStyle) {
        Row row0 = sheet.getRow(0);
        if (row0 == null) {
            row0 = sheet.createRow(0);
        }

        boolean hasHeader = false;
        Cell c0 = row0.getCell(0);
        if (c0 != null && c0.getCellType() == CellType.STRING) {
            hasHeader = "月份".equals(c0.getStringCellValue());
        }
        if (hasHeader) {
            return;
        }

        writeCell(row0, 0, "月份", headerStyle);
        writeCell(row0, 1, "上游客户", headerStyle);
        writeCell(row0, 2, "客户", headerStyle);
        writeCell(row0, 3, "下游客户", headerStyle);
        writeCell(row0, 4, "产品名", headerStyle);
    }

    private int findNextAppendRow(Sheet sheet) {
        int last = sheet.getLastRowNum();
        for (int r = last; r >= 0; r--) {
            Row row = sheet.getRow(r);
            if (!isRowEmpty(row)) {
                return r + 1;
            }
        }
        return 0;
    }

    private boolean isRowEmpty(Row row) {
        if (row == null) {
            return true;
        }
        short last = row.getLastCellNum();
        if (last < 0) {
            return true;
        }
        for (int c = 0; c < last; c++) {
            Cell cell = row.getCell(c);
            if (cell == null) {
                continue;
            }
            if (cell.getCellType() == CellType.BLANK) {
                continue;
            }
            String val = getCellStringValue(cell).trim();
            if (!val.isEmpty()) {
                return false;
            }
        }
        return true;
    }

    private int appendParsedRows(Sheet targetSheet, int rowStart, ParseResult parsed, CellStyle textStyle) {
        int rowCursor = rowStart;
        for (String company : parsed.companyOrder) {
            List<String> upstreams = new ArrayList<>(parsed.upstreams.getOrDefault(company, Collections.emptySet()));
            List<String> downstreams = new ArrayList<>(parsed.downstreams.getOrDefault(company, Collections.emptySet()));

            int rowCount = Math.max(upstreams.size(), downstreams.size());
            if (rowCount == 0) {
                rowCount = 1;
            }

            for (int i = 0; i < rowCount; i++) {
                Row row = targetSheet.createRow(rowCursor++);
                writeCell(row, 0, parsed.monthValue, textStyle);
                writeCell(row, 1, pickByRepeatLast(upstreams, i), textStyle);
                writeCell(row, 2, company, textStyle);
                writeCell(row, 3, pickByRepeatLast(downstreams, i), textStyle);
                writeCell(row, 4, parsed.productValue, textStyle);
            }
        }
        return rowCursor;
    }

    private ParseResult parseSheet(Sheet sheet) {
        ParseResult result = new ParseResult();
        if (sheet.getPhysicalNumberOfRows() < 2) {
            return result;
        }

        Row headerRow = sheet.getRow(0);
        if (headerRow == null) {
            return result;
        }

        List<String> headers = readHeaders(headerRow);
        int supplierIdx = findHeaderIndex(headers, Arrays.asList("供方", "卖方", "供方/发货单位", "供方／发货单位"));
        int demanderIdx = findHeaderIndex(headers, Arrays.asList("需方", "买方", "需方/提货单位", "需方／提货单位"));
        int signingTimeIdx = findHeaderIndex(headers, Collections.singletonList("签约时间"));
        int productIdx = findHeaderIndex(headers, Arrays.asList("产品名", "品名", "商品名称"));

        if (supplierIdx < 0 || demanderIdx < 0) {
            log.warn("Sheet [{}] 缺少必要列（供方/需方），headers={}", sheet.getSheetName(), headers);
            return result;
        }

        Set<String> edgeSet = new LinkedHashSet<>();
        for (int r = 1; r <= sheet.getLastRowNum(); r++) {
            Row row = sheet.getRow(r);
            if (row == null) {
                continue;
            }

            String supplier = getCellStringValue(row.getCell(supplierIdx)).trim();
            String demander = getCellStringValue(row.getCell(demanderIdx)).trim();
            if (supplier.isEmpty() || demander.isEmpty()) {
                continue;
            }

            if (result.monthValue.isEmpty() && signingTimeIdx >= 0) {
                result.monthValue = formatYearMonthFromSigningTime(row.getCell(signingTimeIdx));
            }
            if (result.productValue.isEmpty() && productIdx >= 0) {
                result.productValue = getCellStringValue(row.getCell(productIdx)).trim();
            }

            String edgeKey = supplier + " -> " + demander;
            if (edgeSet.add(edgeKey)) {
                result.edges.add(new Edge(supplier, demander));
            }
        }

        for (Edge edge : result.edges) {
            ensureCompany(result, edge.supplier);
            ensureCompany(result, edge.demander);
            result.downstreams.computeIfAbsent(edge.supplier, k -> new LinkedHashSet<>()).add(edge.demander);
            result.upstreams.computeIfAbsent(edge.demander, k -> new LinkedHashSet<>()).add(edge.supplier);
        }

        return result;
    }

    private void ensureCompany(ParseResult result, String company) {
        if (!result.companySeen.add(company)) {
            return;
        }
        result.companyOrder.add(company);
    }

    private List<String> readHeaders(Row headerRow) {
        List<String> headers = new ArrayList<>();
        short lastCellNum = headerRow.getLastCellNum();
        if (lastCellNum < 0) {
            return headers;
        }
        for (int c = 0; c < lastCellNum; c++) {
            headers.add(getCellStringValue(headerRow.getCell(c)).trim());
        }
        return headers;
    }

    private int findHeaderIndex(List<String> headers, List<String> candidates) {
        for (int i = 0; i < headers.size(); i++) {
            String h = headers.get(i);
            for (String c : candidates) {
                if (h.equals(c)) {
                    return i;
                }
            }
        }
        return -1;
    }

    private String pickByRepeatLast(List<String> values, int index) {
        if (values.isEmpty()) {
            return "";
        }
        if (index < values.size()) {
            return values.get(index);
        }
        return values.get(values.size() - 1);
    }

    private void writeCell(Row row, int col, String val, CellStyle style) {
        Cell cell = row.getCell(col);
        if (cell == null) {
            cell = row.createCell(col, CellType.STRING);
        } else {
            cell.setCellType(CellType.STRING);
        }
        cell.setCellStyle(style);
        cell.setCellValue(val == null ? "" : val);
    }

    private CellStyle createHeaderStyle(Workbook wb) {
        CellStyle style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.LEFT);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        Font font = wb.createFont();
        font.setBold(true);
        style.setFont(font);
        return style;
    }

    private CellStyle createTextStyle(Workbook wb) {
        CellStyle style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.LEFT);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        return style;
    }

    private Workbook openWorkbook(InputStream is, Path file) throws IOException {
        String name = file.getFileName().toString().toLowerCase();
        return name.endsWith(".xlsx") ? new XSSFWorkbook(is) : new HSSFWorkbook(is);
    }

    private String sanitizeSheetName(String name) {
        if (name == null) {
            return "Sheet1";
        }
        String safe = name.replaceAll("[\\\\/:*?\"<>|]", "_").trim();
        if (safe.isEmpty()) {
            safe = "Sheet1";
        }
        return safe.length() <= 31 ? safe : safe.substring(0, 31);
    }

    private String getCellStringValue(Cell cell) {
        if (cell == null) {
            return "";
        }
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    try {
                        return cell.getLocalDateTimeCellValue().toLocalDate().toString();
                    } catch (Exception ignored) {
                        return "";
                    }
                }
                double n = cell.getNumericCellValue();
                if (n == Math.floor(n) && Math.abs(n) < 1e15) {
                    return String.valueOf((long) n);
                }
                return String.valueOf(n);
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                try {
                    switch (cell.getCachedFormulaResultType()) {
                        case STRING:
                            return cell.getStringCellValue();
                        case NUMERIC:
                            if (DateUtil.isCellDateFormatted(cell)) {
                                return cell.getLocalDateTimeCellValue().toLocalDate().toString();
                            }
                            double n2 = cell.getNumericCellValue();
                            if (n2 == Math.floor(n2) && Math.abs(n2) < 1e15) {
                                return String.valueOf((long) n2);
                            }
                            return String.valueOf(n2);
                        case BOOLEAN:
                            return String.valueOf(cell.getBooleanCellValue());
                        default:
                            return "";
                    }
                } catch (Exception ignored) {
                    try {
                        return cell.getStringCellValue();
                    } catch (Exception ignored2) {
                        return cell.getCellFormula();
                    }
                }
            default:
                return "";
        }
    }

    private String formatYearMonthFromSigningTime(Cell cell) {
        if (cell == null) {
            return "";
        }
        try {
            switch (cell.getCellType()) {
                case STRING:
                    return formatYearMonthFromText(cell.getStringCellValue());
                case NUMERIC:
                    return formatYearMonthFromNumeric(cell.getNumericCellValue());
                case FORMULA:
                    switch (cell.getCachedFormulaResultType()) {
                        case STRING:
                            return formatYearMonthFromText(cell.getStringCellValue());
                        case NUMERIC:
                            return formatYearMonthFromNumeric(cell.getNumericCellValue());
                        default:
                            return "";
                    }
                default:
                    return "";
            }
        } catch (Exception ignored) {
            return "";
        }
    }

    private String formatYearMonthFromNumeric(double value) {
        if (Double.isNaN(value) || Double.isInfinite(value)) {
            return "";
        }
        if (value > 20000 && value < 80000) {
            Date date = DateUtil.getJavaDate(value, false);
            LocalDate localDate = Instant.ofEpochMilli(date.getTime())
                    .atZone(ZoneId.systemDefault())
                    .toLocalDate();
            return YEAR_MONTH_FMT.format(localDate);
        }
        String raw = (value == Math.floor(value)) ? String.valueOf((long) value) : String.valueOf(value);
        return formatYearMonthFromText(raw);
    }

    private String formatYearMonthFromText(String raw) {
        if (raw == null) {
            return "";
        }
        String s = raw.trim();
        if (s.isEmpty()) {
            return "";
        }

        if (s.matches("^\\d+(\\.0+)?$")) {
            try {
                double serial = Double.parseDouble(s);
                if (serial > 20000 && serial < 80000) {
                    return formatYearMonthFromNumeric(serial);
                }
            } catch (Exception ignored) {
            }
        }

        String normalized = s.replace("/", "-").replace(".", "-");
        for (DateTimeFormatter fmt : Arrays.asList(
                DateTimeFormatter.ofPattern("yyyy-M-d"),
                DateTimeFormatter.ofPattern("yyyy-MM-dd")
        )) {
            try {
                return YEAR_MONTH_FMT.format(LocalDate.parse(normalized, fmt));
            } catch (Exception ignored) {
            }
        }

        if (normalized.matches("^\\d{8}$")) {
            try {
                return YEAR_MONTH_FMT.format(LocalDate.parse(normalized, DateTimeFormatter.BASIC_ISO_DATE));
            } catch (Exception ignored) {
            }
        }

        if (normalized.matches("^\\d{4}-\\d{1,2}$")) {
            try {
                String[] parts = normalized.split("-");
                int year = Integer.parseInt(parts[0]);
                int month = Integer.parseInt(parts[1]);
                return YEAR_MONTH_FMT.format(LocalDate.of(year, month, 1));
            } catch (Exception ignored) {
            }
        }

        return "";
    }

    private static class Edge {
        final String supplier;
        final String demander;

        Edge(String supplier, String demander) {
            this.supplier = supplier;
            this.demander = demander;
        }
    }

    private static class ParseResult {
        final List<Edge> edges = new ArrayList<>();
        final Set<String> companySeen = new LinkedHashSet<>();
        final List<String> companyOrder = new ArrayList<>();
        final Map<String, Set<String>> upstreams = new LinkedHashMap<>();
        final Map<String, Set<String>> downstreams = new LinkedHashMap<>();
        String monthValue = "";
        String productValue = "";
    }
}
