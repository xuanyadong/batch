package com.tjexcel.batch.service;

import com.tjexcel.batch.config.CustomerBelongingDataExtractConfig;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
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
import java.util.ArrayList;
import java.util.Arrays;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;
import java.util.HashMap;

@Service
public class CustomerBelongingDataExtractService {

    private static final Logger log = LoggerFactory.getLogger(CustomerBelongingDataExtractService.class);
    private static final String SUPPLIER_COL = "供方";
    private static final String DEMANDER_COL = "需方";
    private static final String PROFIT_COL = "账面利润";

    private final CustomerBelongingDataExtractConfig config;
    private final DataFormatter dataFormatter = new DataFormatter();

    public CustomerBelongingDataExtractService(CustomerBelongingDataExtractConfig config) {
        this.config = config;
    }

    /**
     * 客属表数据提取（功能7）：
     * 1) 从客属表按sheet读取公司
     * 2) 在数据表中匹配供方/需方
     * 3) 每个sheet生成文件夹，每个公司输出一个“公司名称-汇总表.xlsx”
     */
    public int extract() throws IOException {
        Path customerPath = Paths.get(config.getCustomerTablePath()).toAbsolutePath();
        Path dataPath = Paths.get(config.getDataTablePath()).toAbsolutePath();
        Path outputDir = Paths.get(config.getOutputDir()).toAbsolutePath();

        if (!Files.exists(customerPath)) {
            throw new FileNotFoundException("客属表不存在: " + customerPath);
        }
        if (!Files.exists(dataPath)) {
            throw new FileNotFoundException("数据表不存在: " + dataPath);
        }
        Files.createDirectories(outputDir);

        List<String> extractColumns = parseExtractColumns(config.getExtractColumns());
        if (extractColumns.isEmpty()) {
            throw new IllegalArgumentException("extract-columns 不能为空");
        }

        Map<String, Set<String>> companiesBySheet = readCompaniesBySheet(customerPath, config.getCompanyColumnName());
        if (companiesBySheet.isEmpty()) {
            log.warn("客属表未读取到任何公司，停止处理。file={}", customerPath);
            return 0;
        }

        SourceData sourceData = readSourceData(dataPath);
        int generatedCount = 0;

        for (Map.Entry<String, Set<String>> sheetCompanies : companiesBySheet.entrySet()) {
            String folderName = sanitizeName(sheetCompanies.getKey(), "sheet");
            Path sheetOutputDir = outputDir.resolve(folderName);
            Files.createDirectories(sheetOutputDir);

            for (String company : sheetCompanies.getValue()) {
                String companyName = defaultIfBlank(company, "").trim();
                if (companyName.isEmpty()) {
                    continue;
                }
                List<Map<String, CellSnapshot>> hitRows = filterRowsForCompany(sourceData.rows, companyName, extractColumns);
                writeCompanyWorkbook(sheetOutputDir, companyName, extractColumns, hitRows, sourceData.columnWidthByHeader);
                generatedCount++;
            }
        }

        log.info("客属表数据提取完成，生成文件 {} 个，输出目录: {}", generatedCount, outputDir);
        return generatedCount;
    }

    private Map<String, Set<String>> readCompaniesBySheet(Path customerPath, String companyColumnName) throws IOException {
        Map<String, Set<String>> result = new LinkedHashMap<>();
        try (InputStream is = Files.newInputStream(customerPath);
             Workbook workbook = openWorkbook(is, customerPath)) {
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                Sheet sheet = workbook.getSheetAt(i);
                if (sheet == null) {
                    continue;
                }
                Set<String> companies = readCompaniesFromSheet(sheet, companyColumnName, evaluator);
                if (!companies.isEmpty()) {
                    result.put(sheet.getSheetName(), companies);
                }
            }
        }
        return result;
    }

    private Set<String> readCompaniesFromSheet(Sheet sheet, String companyColumnName, FormulaEvaluator evaluator) {
        Set<String> companies = new LinkedHashSet<>();
        if (sheet.getPhysicalNumberOfRows() < 2) {
            return companies;
        }
        Row header = sheet.getRow(0);
        int companyCol = findColumnIndex(header, defaultIfBlank(companyColumnName, "企业抬头"));
        if (companyCol < 0) {
            log.warn("sheet={} 未找到公司列: {}", sheet.getSheetName(), companyColumnName);
            return companies;
        }
        for (int r = 1; r <= sheet.getLastRowNum(); r++) {
            Row row = sheet.getRow(r);
            if (row == null) {
                continue;
            }
            String company = getCellText(row.getCell(companyCol), evaluator).trim();
            if (!company.isEmpty()) {
                companies.add(company);
            }
        }
        return companies;
    }

    private SourceData readSourceData(Path dataPath) throws IOException {
        List<Map<String, CellSnapshot>> rows = new ArrayList<>();
        Map<String, Integer> columnWidthByHeader = new LinkedHashMap<>();
        try (InputStream is = Files.newInputStream(dataPath);
             Workbook wb = WorkbookFactory.create(is)) {
            FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
            for (int si = 0; si < wb.getNumberOfSheets(); si++) {
                Sheet sheet = wb.getSheetAt(si);
                if (sheet == null || sheet.getPhysicalNumberOfRows() < 2) {
                    continue;
                }
                Row header = sheet.getRow(0);
                if (header == null) {
                    continue;
                }
                Map<Integer, String> headerByIndex = readHeaderMap(header, evaluator);
                mergeColumnWidths(sheet, headerByIndex, columnWidthByHeader);
                for (int r = 1; r <= sheet.getLastRowNum(); r++) {
                    Row row = sheet.getRow(r);
                    if (row == null) {
                        continue;
                    }
                    Map<String, CellSnapshot> rowMap = new LinkedHashMap<>();
                    for (Map.Entry<Integer, String> entry : headerByIndex.entrySet()) {
                        Cell sourceCell = row.getCell(entry.getKey());
                        rowMap.put(entry.getValue(), toCellSnapshot(sourceCell, evaluator));
                    }
                    if (!isEmptyDataRow(rowMap)) {
                        rows.add(rowMap);
                    }
                }
            }
        }
        return new SourceData(rows, columnWidthByHeader);
    }

    private void mergeColumnWidths(Sheet sheet,
                                   Map<Integer, String> headerByIndex,
                                   Map<String, Integer> columnWidthByHeader) {
        for (Map.Entry<Integer, String> entry : headerByIndex.entrySet()) {
            int width = sheet.getColumnWidth(entry.getKey());
            String headerName = entry.getValue();
            Integer existing = columnWidthByHeader.get(headerName);
            if (existing == null || width > existing) {
                columnWidthByHeader.put(headerName, width);
            }
        }
    }

    private Map<Integer, String> readHeaderMap(Row header, FormulaEvaluator evaluator) {
        Map<Integer, String> result = new LinkedHashMap<>();
        short last = header.getLastCellNum();
        if (last < 0) {
            return result;
        }
        for (int c = 0; c < last; c++) {
            String name = getCellText(header.getCell(c), evaluator).trim();
            if (!name.isEmpty()) {
                result.put(c, name);
            }
        }
        return result;
    }

    private List<Map<String, CellSnapshot>> filterRowsForCompany(List<Map<String, CellSnapshot>> rows,
                                                           String company,
                                                           List<String> extractColumns) {
        List<Map<String, CellSnapshot>> result = new ArrayList<>();
        for (Map<String, CellSnapshot> row : rows) {
            String supplier = defaultIfBlank(asText(row.get(SUPPLIER_COL)), "").trim();
            String demander = defaultIfBlank(asText(row.get(DEMANDER_COL)), "").trim();
            boolean isSupplier = company.equals(supplier);
            boolean isDemander = company.equals(demander);
            if (!isSupplier && !isDemander) {
                continue;
            }

            Map<String, CellSnapshot> out = new LinkedHashMap<>();
            for (String col : extractColumns) {
                if (PROFIT_COL.equals(col) && isDemander && !isSupplier) {
                    out.put(col, CellSnapshot.blank());
                } else {
                    out.put(col, row.getOrDefault(col, CellSnapshot.blank()));
                }
            }
            result.add(out);
        }
        return result;
    }

    private void writeCompanyWorkbook(Path sheetOutputDir,
                                      String companyName,
                                      List<String> extractColumns,
                                      List<Map<String, CellSnapshot>> rows,
                                      Map<String, Integer> sourceColumnWidthByHeader) throws IOException {
        Path file = sheetOutputDir.resolve(sanitizeName(companyName, "company") + "-汇总表.xlsx");
        try (Workbook wb = new XSSFWorkbook()) {
            Sheet outSheet = wb.createSheet("汇总");
            Map<String, CellStyle> dataFormatStyleCache = new HashMap<>();
            CreationHelper creationHelper = wb.getCreationHelper();
            Row header = outSheet.createRow(0);
            for (int c = 0; c < extractColumns.size(); c++) {
                Cell cell = header.createCell(c, CellType.STRING);
                cell.setCellValue(extractColumns.get(c));
            }
            int rowIndex = 1;
            for (Map<String, CellSnapshot> rowData : rows) {
                Row row = outSheet.createRow(rowIndex++);
                for (int c = 0; c < extractColumns.size(); c++) {
                    String col = extractColumns.get(c);
                    CellSnapshot snapshot = rowData.getOrDefault(col, CellSnapshot.blank());
                    Cell cell = row.createCell(c, snapshot.outputType);
                    applySnapshot(cell, snapshot);
                    applyDataFormatStyle(cell, snapshot, wb, creationHelper, dataFormatStyleCache);
                }
            }
            for (int c = 0; c < extractColumns.size(); c++) {
                String colName = extractColumns.get(c);
                Integer sourceWidth = sourceColumnWidthByHeader.get(colName);
                if (sourceWidth != null && sourceWidth > 0) {
                    outSheet.setColumnWidth(c, sourceWidth);
                } else {
                    outSheet.autoSizeColumn(c);
                }
            }
            try (OutputStream os = Files.newOutputStream(file)) {
                wb.write(os);
            }
        }
    }

    private int findColumnIndex(Row header, String columnName) {
        if (header == null) {
            return -1;
        }
        short last = header.getLastCellNum();
        if (last < 0) {
            return -1;
        }
        String expected = defaultIfBlank(columnName, "").trim();
        for (int c = 0; c < last; c++) {
            String cell = getCellText(header.getCell(c), null).trim();
            if (expected.equals(cell)) {
                return c;
            }
        }
        return -1;
    }

    private String getCellText(Cell cell, FormulaEvaluator evaluator) {
        if (cell == null) {
            return "";
        }
        try {
            short formatIndex = cell.getCellStyle() == null ? 0 : cell.getCellStyle().getDataFormat();
            String formatString = cell.getCellStyle() == null ? null : cell.getCellStyle().getDataFormatString();

            if (cell.getCellType() == CellType.FORMULA && evaluator != null) {
                CellValue evaluated = evaluator.evaluate(cell);
                if (evaluated == null) {
                    return "";
                }
                switch (evaluated.getCellType()) {
                    case STRING:
                        return evaluated.getStringValue();
                    case NUMERIC:
                        return dataFormatter.formatRawCellContents(
                                evaluated.getNumberValue(),
                                formatIndex,
                                formatString
                        );
                    case BOOLEAN:
                        return String.valueOf(evaluated.getBooleanValue());
                    default:
                        return "";
                }
            }

            if (cell.getCellType() == CellType.NUMERIC) {
                return dataFormatter.formatRawCellContents(
                        cell.getNumericCellValue(),
                        formatIndex,
                        formatString
                );
            }
            return dataFormatter.formatCellValue(cell);
        } catch (Exception e) {
            return "";
        }
    }

    private List<String> parseExtractColumns(String configured) {
        String raw = defaultIfBlank(configured, "");
        if (raw.isEmpty()) {
            return new ArrayList<>();
        }
        List<String> cols = new ArrayList<>();
        Arrays.stream(raw.split("\\|"))
                .map(String::trim)
                .filter(s -> !s.isEmpty())
                .forEach(cols::add);
        return cols;
    }

    private boolean isEmptyDataRow(Map<String, CellSnapshot> rowMap) {
        for (CellSnapshot value : rowMap.values()) {
            if (value != null && !defaultIfBlank(value.textValue, "").trim().isEmpty()) {
                return false;
            }
        }
        return true;
    }

    private CellSnapshot toCellSnapshot(Cell cell, FormulaEvaluator evaluator) {
        if (cell == null) {
            return CellSnapshot.blank();
        }
        String text = getCellText(cell, evaluator);
        String dataFormatString = "";
        short dataFormatIndex = 0;
        if (cell.getCellStyle() != null) {
            dataFormatString = defaultIfBlank(cell.getCellStyle().getDataFormatString(), "");
            dataFormatIndex = cell.getCellStyle().getDataFormat();
        }
        try {
            if (cell.getCellType() == CellType.FORMULA && evaluator != null) {
                CellValue evaluated = evaluator.evaluate(cell);
                if (evaluated == null) {
                    return new CellSnapshot(CellType.STRING, text, 0d, false, dataFormatString, dataFormatIndex);
                }
                switch (evaluated.getCellType()) {
                    case NUMERIC:
                        return new CellSnapshot(CellType.NUMERIC, text, evaluated.getNumberValue(), false, dataFormatString, dataFormatIndex);
                    case BOOLEAN:
                        return new CellSnapshot(CellType.BOOLEAN, text, 0d, evaluated.getBooleanValue(), dataFormatString, dataFormatIndex);
                    case STRING:
                    default:
                        return new CellSnapshot(CellType.STRING, text, 0d, false, dataFormatString, dataFormatIndex);
                }
            }
            if (cell.getCellType() == CellType.NUMERIC) {
                return new CellSnapshot(CellType.NUMERIC, text, cell.getNumericCellValue(), false, dataFormatString, dataFormatIndex);
            }
            if (cell.getCellType() == CellType.BOOLEAN) {
                return new CellSnapshot(CellType.BOOLEAN, text, 0d, cell.getBooleanCellValue(), dataFormatString, dataFormatIndex);
            }
        } catch (Exception ignored) {
        }
        return new CellSnapshot(CellType.STRING, text, 0d, false, dataFormatString, dataFormatIndex);
    }

    private String asText(CellSnapshot snapshot) {
        if (snapshot == null) {
            return "";
        }
        return defaultIfBlank(snapshot.textValue, "");
    }

    private void applySnapshot(Cell cell, CellSnapshot snapshot) {
        if (snapshot == null) {
            cell.setCellType(CellType.STRING);
            cell.setCellValue("");
            return;
        }
        switch (snapshot.outputType) {
            case NUMERIC:
                cell.setCellValue(snapshot.numericValue);
                break;
            case BOOLEAN:
                cell.setCellValue(snapshot.booleanValue);
                break;
            default:
                cell.setCellType(CellType.STRING);
                cell.setCellValue(defaultIfBlank(snapshot.textValue, ""));
                break;
        }
    }

    private void applyDataFormatStyle(Cell cell,
                                      CellSnapshot snapshot,
                                      Workbook wb,
                                      CreationHelper creationHelper,
                                      Map<String, CellStyle> styleCache) {
        if (snapshot == null || snapshot.outputType == CellType.STRING) {
            return;
        }
        String fmt = defaultIfBlank(snapshot.dataFormatString, "");
        if (fmt.isEmpty()) {
            return;
        }
        String key = snapshot.outputType.name() + "|" + fmt;
        CellStyle style = styleCache.get(key);
        if (style == null) {
            style = wb.createCellStyle();
            style.setDataFormat(creationHelper.createDataFormat().getFormat(fmt));
            styleCache.put(key, style);
        }
        cell.setCellStyle(style);
    }

    private Workbook openWorkbook(InputStream is, Path file) throws IOException {
        String name = file.getFileName().toString().toLowerCase(Locale.ROOT);
        return name.endsWith(".xlsx") ? new XSSFWorkbook(is) : new HSSFWorkbook(is);
    }

    private String sanitizeName(String raw, String fallback) {
        String safe = defaultIfBlank(raw, fallback).replaceAll("[\\\\/:*?\"<>|]", "_").trim();
        return safe.isEmpty() ? fallback : safe;
    }

    private String defaultIfBlank(String value, String defaultValue) {
        if (value == null || value.trim().isEmpty()) {
            return defaultValue;
        }
        return value.trim();
    }

    private static class SourceData {
        final List<Map<String, CellSnapshot>> rows;
        final Map<String, Integer> columnWidthByHeader;

        SourceData(List<Map<String, CellSnapshot>> rows, Map<String, Integer> columnWidthByHeader) {
            this.rows = rows;
            this.columnWidthByHeader = columnWidthByHeader;
        }
    }

    private static class CellSnapshot {
        final CellType outputType;
        final String textValue;
        final double numericValue;
        final boolean booleanValue;
        final String dataFormatString;
        final short dataFormatIndex;

        CellSnapshot(CellType outputType,
                     String textValue,
                     double numericValue,
                     boolean booleanValue,
                     String dataFormatString,
                     short dataFormatIndex) {
            this.outputType = outputType;
            this.textValue = textValue;
            this.numericValue = numericValue;
            this.booleanValue = booleanValue;
            this.dataFormatString = dataFormatString;
            this.dataFormatIndex = dataFormatIndex;
        }

        static CellSnapshot blank() {
            return new CellSnapshot(CellType.STRING, "", 0d, false, "", (short) 0);
        }
    }
}
