package com.tjexcel.batch.service;

import com.tjexcel.batch.config.ContractConfig;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;

import net.sourceforge.pinyin4j.PinyinHelper;
import net.sourceforge.pinyin4j.format.HanyuPinyinOutputFormat;
import net.sourceforge.pinyin4j.format.HanyuPinyinToneType;

import java.io.*;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDate;
import java.time.YearMonth;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 合同批量生成服务
 * 从数据表读取每一行，填充模板中的占位符，批量生成合同文件
 */
@Service
public class ContractGeneratorService {

    private static final Logger log = LoggerFactory.getLogger(ContractGeneratorService.class);

    private final ContractConfig config;

    public ContractGeneratorService(ContractConfig config) {
        this.config = config;
    }

    /**
     * 执行批量生成
     */
    public int generate() throws IOException {
        String dataPath = config.getDataPath();
        String templatePath = config.getTemplatePath();
        String outputDir = config.getOutputDir();

        // 解析为绝对路径
        Path dataFile = Paths.get(dataPath).toAbsolutePath();
        Path templateFile = Paths.get(templatePath).toAbsolutePath();
        Path outputPath = Paths.get(outputDir).toAbsolutePath();

        if (!Files.exists(dataFile)) {
            throw new FileNotFoundException("数据表不存在: " + dataFile);
        }
        if (!Files.exists(templateFile)) {
            throw new FileNotFoundException("模板文件不存在: " + templateFile);
        }

        Files.createDirectories(outputPath);

        List<Map<String, String>> rows = readDataSheet(dataFile);
        log.info("从 {} 读取到 {} 行数据", dataFile, rows.size());
        if (!rows.isEmpty()) {
            log.info("数据表列名（可用于 output-file-name-pattern）: {}", rows.get(0).keySet());
        }

        if (rows.isEmpty()) {
            log.warn("数据表为空，无合同可生成");
            return 0;
        }

        Map<String, Integer> contractSeqMap = new HashMap<>();
        int successCount = 0;
        Set<String> usedFileNames = new HashSet<>();
        for (int i = 0; i < rows.size(); i++) {
            Map<String, String> rowData = new LinkedHashMap<>(rows.get(i));
            String contractNo = generateContractNumber(rowData, contractSeqMap);
            rowData.put("合同编号", contractNo);

            try {
                String outputFileName = resolveFileName(rowData, config.getOutputFileNamePattern());
                outputFileName = ensureUniqueFileName(outputFileName, usedFileNames, i + 1);
                usedFileNames.add(outputFileName);
                Path outputFile = outputPath.resolve(outputFileName);
                fillAndSave(templateFile, rowData, outputFile);
                log.info("[{}/{}] 已生成: {}", i + 1, rows.size(), outputFileName);
                successCount++;
            } catch (Exception e) {
                log.error("第 {} 行生成失败: {}", i + 2, e.getMessage(), e);
            }
        }

        log.info("批量生成完成，成功 {}/{} 个合同，输出目录: {}", successCount, rows.size(), outputPath);
        return successCount;
    }

    /**
     * 合同编号规则：供方简称首字母-需方简称首字母-日期(yyyyMMdd)-序号(001/002...)
     * 同一供方简称+需方简称+日期 下序号递增，如 YK-QY-20260310-001、YK-QY-20260310-002
     */
    private String generateContractNumber(Map<String, String> rowData, Map<String, Integer> seqMap) {
        String gongFang = emptyToBlank(rowData.get("供方简称"));
        String xuFang = emptyToBlank(rowData.get("需方简称"));
        String dateStr = emptyToBlank(rowData.get("签约时间"));
        String gongInitials = toPinyinInitials(gongFang).toUpperCase();
        String xuInitials = toPinyinInitials(xuFang).toUpperCase();
        if (gongInitials.isEmpty()) gongInitials = "X";
        if (xuInitials.isEmpty()) xuInitials = "X";
        String yyyyMMdd = formatContractDate(dateStr);
        String groupKey = gongFang + "|" + xuFang + "|" + yyyyMMdd;
        int seq = seqMap.merge(groupKey, 1, Integer::sum);
        return gongInitials + "-" + xuInitials + "-" + yyyyMMdd + "-" + String.format("%03d", seq);
    }

    private String emptyToBlank(String s) {
        return s == null ? "" : s.trim();
    }

    private String formatContractDate(String dateStr) {
        if (dateStr == null || dateStr.trim().isEmpty()) return LocalDate.now().format(DateTimeFormatter.BASIC_ISO_DATE);
        String s = dateStr.trim().replace(" ", "").replace("/", "-");
        DateTimeFormatter[] formatters = {
            DateTimeFormatter.ofPattern("yyyy-M-d"),
            DateTimeFormatter.ofPattern("yyyy-MM-dd"),
            DateTimeFormatter.ISO_LOCAL_DATE
        };
        for (DateTimeFormatter f : formatters) {
            try {
                return LocalDate.parse(s, f).format(DateTimeFormatter.BASIC_ISO_DATE);
            } catch (DateTimeParseException ignored) {}
        }
        return LocalDate.now().format(DateTimeFormatter.BASIC_ISO_DATE);
    }

    /** 汉字转拼音首字母，如 伊科 -> YK */
    private String toPinyinInitials(String str) {
        if (str == null || str.isEmpty()) return "";
        HanyuPinyinOutputFormat format = new HanyuPinyinOutputFormat();
        format.setToneType(HanyuPinyinToneType.WITHOUT_TONE);
        StringBuilder sb = new StringBuilder();
        for (char c : str.toCharArray()) {
            if (Character.toString(c).matches("[\\u4e00-\\u9fa5]")) {
                try {
                    String[] py = PinyinHelper.toHanyuPinyinStringArray(c, format);
                    if (py != null && py.length > 0) sb.append(py[0].charAt(0));
                } catch (Exception ignored) {}
            } else if (Character.isLetterOrDigit(c)) {
                sb.append(c);
            }
        }
        return sb.toString();
    }

    /** 按文件内容判断格式：xlsx 以 PK 开头，xls 为 OLE2 */
    private Workbook openWorkbook(InputStream is, Path file) throws IOException {
        PushbackInputStream pis = new PushbackInputStream(is, 8);
        byte[] header = new byte[4];
        int n = pis.read(header);
        if (n > 0) pis.unread(header, 0, n);
        boolean isXlsx = n >= 2 && header[0] == 0x50 && header[1] == 0x4B;
        return isXlsx ? new XSSFWorkbook(pis) : new HSSFWorkbook(pis);
    }

    /**
     * 读取数据表第一个 sheet，第一行为表头
     */
    private List<Map<String, String>> readDataSheet(Path dataFile) throws IOException {
        List<Map<String, String>> rows = new ArrayList<>();

        try (InputStream is = Files.newInputStream(dataFile);
             Workbook workbook = openWorkbook(is, dataFile)) {

            Sheet sheet = workbook.getSheetAt(0);
            if (sheet == null || sheet.getPhysicalNumberOfRows() < 2) {
                return rows;
            }

            Row headerRow = sheet.getRow(0);
            List<String> headers = new ArrayList<>();
            for (Cell cell : headerRow) {
                headers.add(getCellStringValue(cell, null));
            }

            for (int r = 1; r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;

                Map<String, String> rowData = new LinkedHashMap<>();
                for (int c = 0; c < headers.size(); c++) {
                    String header = headers.get(c);
                    if (header == null || header.trim().isEmpty()) {
                        header = "col" + c;
                    }
                    Cell cell = row.getCell(c);
                    String colName = header.trim();
                    rowData.put(colName, cell != null ? getCellStringValue(cell, colName) : "");
                }
                rows.add(rowData);
            }
        }
        return rows;
    }

    /**
     * 填充模板并保存
     */
    private void fillAndSave(Path templateFile, Map<String, String> rowData, Path outputFile) throws IOException {
        try (InputStream is = Files.newInputStream(templateFile)) {
            Workbook workbook = openWorkbook(is, templateFile);

            replacePlaceholders(workbook, rowData);

            int minWidthChars = computeMinNumericColumnWidth(rowData);
            autoSizeColumns(workbook, minWidthChars);

            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            evaluator.evaluateAll();

            if (workbook instanceof org.apache.poi.xssf.usermodel.XSSFWorkbook) {
                ((org.apache.poi.xssf.usermodel.XSSFWorkbook) workbook).setForceFormulaRecalculation(true);
            } else if (workbook instanceof HSSFWorkbook) {
                ((HSSFWorkbook) workbook).setForceFormulaRecalculation(true);
            }

            try (OutputStream os = Files.newOutputStream(outputFile)) {
                workbook.write(os);
            }
            workbook.close();
        }
    }

    /**
     * 合并列名和列字母(A,B,C)到替换 Map，支持 ${列名} 和 ${A}
     */
    private Map<String, String> buildReplacementMap(Map<String, String> rowData) {
        Map<String, String> map = new HashMap<>(rowData);
        List<String> values = new ArrayList<>(rowData.values());
        for (int i = 0; i < values.size() && i < 26; i++) {
            map.put(String.valueOf((char) ('A' + i)), values.get(i) != null ? values.get(i) : "");
        }
        map.put("date", YearMonth.now().atEndOfMonth().format(DateTimeFormatter.ofPattern("yyyy年MM月dd日")));
        return map;
    }

    /** 价税合计金额、增值税额、不含税金额三列中最长显示长度，且不小于 22（预留千分位等空间） */
    private int computeMinNumericColumnWidth(Map<String, String> rowData) {
        int minWidth = 22;
        for (String key : Arrays.asList("价税合计金额", "增值税额", "不含税金额")) {
            String v = rowData.get(key);
            if (v != null && !v.isEmpty()) {
                minWidth = Math.max(minWidth, v.length() + 4);
            }
        }
        return minWidth;
    }

    /** 根据内容自动调整列宽；仅对 C/F/G/H 四列套用最小宽度，避免金额列 ######## */
    private void autoSizeColumns(Workbook workbook, int minWidthChars) {
        int minWidthUnits = minWidthChars * 256;
        for (int s = 0; s < workbook.getNumberOfSheets(); s++) {
            Sheet sheet = workbook.getSheetAt(s);
            int maxCol = 0;
            for (Row row : sheet) {
                if (row != null && row.getLastCellNum() > maxCol) {
                    maxCol = row.getLastCellNum();
                }
            }
            for (int c = 0; c < maxCol; c++) {
                // 仅对 C/F/G/H 四列（索引 2/5/6/7）做自动列宽和最小宽度，其它列完全按模板保持不动
                if (c == 2 || c == 5 || c == 6 || c == 7) {
                    try {
                        sheet.autoSizeColumn(c);
                        int w = sheet.getColumnWidth(c);
                        if (w < minWidthUnits) {
                            sheet.setColumnWidth(c, minWidthUnits);
                        }
                    } catch (Exception ignored) {}
                }
            }
        }
    }

    /**
     * 替换工作簿中占位符 ${列名}，列名与数据表表头一致；空值填空白；不含占位符的公式保留。${date} 为当月最后一天，格式 yyyy年MM月dd日。
     */
    private void replacePlaceholders(Workbook workbook, Map<String, String> rowData) {
        Map<String, String> replaceMap = buildReplacementMap(rowData);
        String prefix = config.getPlaceholderPrefix();
        String suffix = config.getPlaceholderSuffix();
        Pattern pattern = Pattern.compile(Pattern.quote(prefix) + "([^" + Pattern.quote(suffix) + "]+)" + Pattern.quote(suffix));

        for (int s = 0; s < workbook.getNumberOfSheets(); s++) {
            Sheet sheet = workbook.getSheetAt(s);
            for (Row row : sheet) {
                for (Cell cell : row) {
                    if (cell.getCellType() == CellType.STRING) {
                        String value = cell.getStringCellValue();
                        String replaced = replaceInString(value, replaceMap, pattern, prefix, suffix);
                        if (!value.equals(replaced)) {
                            setCellValueByType(cell, replaced);
                        }
                    } else if (cell.getCellType() == CellType.FORMULA) {
                        // 仅当公式中含占位符时才修改，否则保留原公式
                        try {
                            String formula = cell.getCellFormula();
                            if (pattern.matcher(formula).find()) {
                                String replaced = replaceInString(formula, replaceMap, pattern, prefix, suffix);
                                if (!formula.equals(replaced)) {
                                    cell.setCellFormula(replaced);
                                }
                            }
                        } catch (Exception ignored) {
                        }
                    }
                }
            }
        }
    }

    private String replaceInString(String str, Map<String, String> replaceMap, Pattern pattern, String prefix, String suffix) {
        if (str == null) return "";
        Matcher matcher = pattern.matcher(str);
        StringBuffer sb = new StringBuffer();
        while (matcher.find()) {
            String key = matcher.group(1).trim();
            String replacement = replaceMap.getOrDefault(key, "");
            matcher.appendReplacement(sb, Matcher.quoteReplacement(replacement));
        }
        matcher.appendTail(sb);
        return sb.toString();
    }

    /**
     * 根据数据解析输出文件名，支持 ${列名} 和 ${A}、${B}
     */
    private String resolveFileName(Map<String, String> rowData, String pattern) {
        Map<String, String> replaceMap = buildReplacementMap(rowData);
        String result = pattern;
        for (Map.Entry<String, String> e : replaceMap.entrySet()) {
            result = result.replace(config.getPlaceholderPrefix() + e.getKey() + config.getPlaceholderSuffix(),
                    e.getValue() != null ? e.getValue() : "");
        }
        result = result.replaceAll(Pattern.quote(config.getPlaceholderPrefix()) + "[^" + Pattern.quote(config.getPlaceholderSuffix()) + "]*" + Pattern.quote(config.getPlaceholderSuffix()), "");
        return result.replaceAll("[\\\\/:*?\"<>|]", "_");
    }

    /**
     * 确保文件名唯一，避免重复时互相覆盖。
     * 当文件名已存在时，在扩展名前插入 _行号
     */
    private String ensureUniqueFileName(String fileName, Set<String> usedFileNames, int rowIndex) {
        if (!usedFileNames.contains(fileName)) {
            return fileName;
        }
        int lastDot = fileName.lastIndexOf('.');
        String base = lastDot > 0 ? fileName.substring(0, lastDot) : fileName;
        String ext = lastDot > 0 ? fileName.substring(lastDot) : "";
        return base + "_" + rowIndex + ext;
    }

    /**
     * 读取单元格为字符串。columnHeader 用于金额列保留两位小数、日期按 yyyy/M/d 输出。
     */
    private String getCellStringValue(Cell cell, String columnHeader) {
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getLocalDateTimeCellValue().toLocalDate().format(DATE_DISPLAY_FMT);
                }
                int decimals = (columnHeader != null && AMOUNT_TWO_DECIMAL_COLUMNS.contains(columnHeader)) ? 2 : -1;
                return formatNumericForDisplay(cell.getNumericCellValue(), decimals);
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                try {
                    double num = cell.getNumericCellValue();
                    int dec = (columnHeader != null && AMOUNT_TWO_DECIMAL_COLUMNS.contains(columnHeader)) ? 2 : -1;
                    return formatNumericForDisplay(num, dec);
                } catch (Exception e) {
                    try {
                        return cell.getStringCellValue();
                    } catch (Exception e2) {
                        return cell.getCellFormula();
                    }
                }
            default:
                return "";
        }
    }

    /** 写入单元格：能解析为数字则写数字（便于 SUM 等公式计算），否则写文本 */
    private void setCellValueByType(Cell cell, String value) {
        if (value == null) value = "";
        value = value.trim();
        try {
            double num = Double.parseDouble(value.replace(",", ""));
            cell.setCellValue(num);
        } catch (NumberFormatException e) {
            cell.setCellValue(value);
        }
    }

    /** 数值转字符串：不用科学计数法。decimalPlaces=-1 表示自动；>=0 表示保留该位小数四舍五入 */
    private static String formatNumericForDisplay(double num, int decimalPlaces) {
        if (Double.isNaN(num) || Double.isInfinite(num)) return String.valueOf(num);
        if (decimalPlaces == 0 || (decimalPlaces < 0 && num == Math.floor(num) && Math.abs(num) < 1e15)) {
            return String.valueOf((long) num);
        }
        int scale = decimalPlaces >= 0 ? decimalPlaces : 6;
        BigDecimal bd = BigDecimal.valueOf(num).setScale(scale, RoundingMode.HALF_UP);
        return bd.stripTrailingZeros().toPlainString();
    }

    private static final Set<String> AMOUNT_TWO_DECIMAL_COLUMNS = new HashSet<>(
            Arrays.asList("价税合计金额", "增值税额", "不含税金额"));

    private static final DateTimeFormatter DATE_DISPLAY_FMT = DateTimeFormatter.ofPattern("yyyy/M/d");
}
