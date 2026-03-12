package com.tjexcel.batch.service;

import com.tjexcel.batch.config.BillOfLadingConfig;
import com.tjexcel.batch.util.OrderSplitUtil;
import net.sourceforge.pinyin4j.PinyinHelper;
import net.sourceforge.pinyin4j.format.HanyuPinyinOutputFormat;
import net.sourceforge.pinyin4j.format.HanyuPinyinToneType;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;

import java.io.*;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 提单批量生成服务
 * 从数据表读取每一行，替换占位符，填充内置表格（数量拆分后逐行）
 * <p>
 * 卡号全局对应：同一公司作为需方收货的卡号总量 = 作为供方发货的卡号总量，
 * 保证收货/发货的卡号能够一一对应。
 */
@Service
public class BillOfLadingService {

    private static final Logger log = LoggerFactory.getLogger(BillOfLadingService.class);

    private static final int PIECES_MULTIPLIER = 40;
    private static final int ALLOCATION_RETRY_LIMIT = 20;

    /** 表格列：产品名 规格型号 卡号 数量（吨） 件数 备注 */
    private static final String[] TABLE_HEADERS = {"产品名", "规格型号", "卡号", "数量（吨）", "件数", "备注"};

    private final BillOfLadingConfig config;

    public BillOfLadingService(BillOfLadingConfig config) {
        this.config = config;
    }

    public int generate() throws IOException {
        Path dataFile = Paths.get(config.getDataPath()).toAbsolutePath();
        Path templateFile = Paths.get(config.getTemplatePath()).toAbsolutePath();
        Path outputPath = Paths.get(config.getOutputDir()).toAbsolutePath();

        if (!Files.exists(dataFile)) {
            throw new FileNotFoundException("提单数据表不存在: " + dataFile);
        }
        if (!Files.exists(templateFile)) {
            throw new FileNotFoundException("提单模板不存在: " + templateFile);
        }

        Files.createDirectories(outputPath);

        List<Map<String, String>> rows = readDataSheet(dataFile);
        log.info("从 {} 读取到 {} 行提单数据", dataFile, rows.size());
        if (rows.isEmpty()) {
            log.warn("提单数据表为空，无提单可生成");
            return 0;
        }

        // 预计算：按 (公司, 规格型号) 建立卡号库存，并预分配每行的卡号明细
        CardAllocationContext ctx = buildCardAllocationContext(rows);

        Map<String, Integer> seqMap = new HashMap<>();
        int successCount = 0;
        Set<String> usedFileNames = new HashSet<>();
        for (int i = 0; i < rows.size(); i++) {
            Map<String, String> rowData = new LinkedHashMap<>(rows.get(i));
            String billNo = generateBillNumber(rowData, seqMap);
            rowData.put("提单编号", billNo);
            rowData.put("供方简称", firstFourChars(emptyToBlank(getColumnValue(rowData, "供方/发货单位", "供方／发货单位"))));
            rowData.put("需方简称", firstFourChars(emptyToBlank(getColumnValue(rowData, "需方/提货单位", "需方／提货单位"))));

            try {
                String outputFileName = resolveFileName(rowData, config.getOutputFileNamePattern());
                outputFileName = ensureUniqueFileName(outputFileName, usedFileNames, i + 1);
                usedFileNames.add(outputFileName);
                Path outputFile = outputPath.resolve(outputFileName);
                fillAndSave(templateFile, rowData, outputFile, ctx, i);
                log.info("[{}/{}] 已生成: {}", i + 1, rows.size(), outputFileName);
                successCount++;
            } catch (Exception e) {
                log.error("第 {} 行提单生成失败: {}", i + 2, e.getMessage(), e);
            }
        }

        log.info("提单批量生成完成，成功 {}/{} 个，输出目录: {}", successCount, rows.size(), outputPath);
        return successCount;
    }

    /**
     * 预计算卡号分配：按 (供方, 规格型号) 建立库存，为每行预分配 (卡号, 数量) 列表
     */
    private CardAllocationContext buildCardAllocationContext(List<Map<String, String>> rows) {
        int min = config.getSplitMin();
        int max = config.getSplitMax();
        int maxSubCount = config.getSplitMaxSubCount();

        // 按 (供方, 规格型号) 分组发货行
        Map<String, List<RowRef>> shipmentByCompanySpec = new LinkedHashMap<>();
        for (int i = 0; i < rows.size(); i++) {
            Map<String, String> row = rows.get(i);
            String gongFang = emptyToBlank(getColumnValue(row, "供方/发货单位", "供方／发货单位"));
            String specModel = emptyToBlank(getColumnValue(row, "规格型号"));
            int qty = (int) Math.round(parseQuantity(row));
            if (gongFang.isEmpty() || qty <= 0) continue;
            String key = gongFang + "|" + specModel;
            shipmentByCompanySpec.computeIfAbsent(key, k -> new ArrayList<>()).add(new RowRef(i, qty, row));
        }

        // 对每个 (供方, 规格型号)：按收货行收集数量，每行若超过 splitMax 则拆成多张卡（每张在 [splitMin, splitMax]），保证收=发且每项满足限制
        Map<String, Map<String, Integer>> inventoryByCompanySpec = new HashMap<>();
        for (String key : shipmentByCompanySpec.keySet()) {
            String[] parts = key.split("\\|", 2);
            String company = parts[0];
            String specModel = parts[1];

            List<Integer> receiptRowQtys = new ArrayList<>();
            for (Map<String, String> row : rows) {
                String xuFang = emptyToBlank(getColumnValue(row, "需方/提货单位", "需方／提货单位"));
                String spec = emptyToBlank(getColumnValue(row, "规格型号"));
                if (xuFang.equals(company) && spec.equals(specModel)) {
                    int q = (int) Math.round(parseQuantity(row));
                    if (q > 0) receiptRowQtys.add(q);
                }
            }

            if (receiptRowQtys.isEmpty()) continue;

            // 每行数量若 > splitMax 则用 OrderSplitUtil 拆成多张卡（每张在 [min,max]）；否则一行一卡
            List<Integer> cardQtys = new ArrayList<>();
            for (Integer rowQty : receiptRowQtys) {
                if (rowQty <= max) {
                    cardQtys.add(rowQty);
                } else {
                    try {
                        cardQtys.addAll(OrderSplitUtil.split(rowQty, min, max, maxSubCount));
                    } catch (IllegalArgumentException e) {
                        log.warn("{} 收货行数量 {} 拆单失败: {}，按单卡处理", key, rowQty, e.getMessage());
                        cardQtys.add(rowQty);
                    }
                }
            }

            Map<String, Integer> inventory = new LinkedHashMap<>();
            for (int j = 0; j < cardQtys.size(); j++) {
                String cardNo = specModel + "-" + (j + 1);
                inventory.put(cardNo, cardQtys.get(j));
            }
            inventoryByCompanySpec.put(key, inventory);
        }

        // 为每行分配 (卡号, 数量)：严格按数据表行顺序处理，从对应 (供方,规格) 的库存中扣减
        List<List<CardQty>> rowAllocations = new ArrayList<>(rows.size());
        for (int i = 0; i < rows.size(); i++) {
            rowAllocations.add(null);
        }

        Map<String, Map<String, Integer>> mutableInventory = new HashMap<>();
        for (Map.Entry<String, Map<String, Integer>> e : inventoryByCompanySpec.entrySet()) {
            Map<String, Integer> copy = new LinkedHashMap<>();
            for (Map.Entry<String, Integer> e2 : e.getValue().entrySet()) {
                copy.put(e2.getKey(), e2.getValue());
            }
            mutableInventory.put(e.getKey(), copy);
        }

        for (int rowIndex = 0; rowIndex < rows.size(); rowIndex++) {
            Map<String, String> row = rows.get(rowIndex);
            String gongFang = emptyToBlank(getColumnValue(row, "供方/发货单位", "供方／发货单位"));
            String specModel = emptyToBlank(getColumnValue(row, "规格型号"));
            int qty = (int) Math.round(parseQuantity(row));
            if (gongFang.isEmpty() || qty <= 0) continue;

            String key = gongFang + "|" + specModel;
            Map<String, Integer> inv = mutableInventory.get(key);
            if (inv == null) continue;

            List<CardQty> allocation = allocateFromInventory(qty, inv, min, max, maxSubCount);
            if (allocation != null) {
                rowAllocations.set(rowIndex, allocation);
                for (CardQty cq : allocation) {
                    int remain = inv.getOrDefault(cq.cardNo, 0) - cq.qty;
                    if (remain <= 0) {
                        inv.remove(cq.cardNo);
                    } else {
                        inv.put(cq.cardNo, remain);
                    }
                }
            } else {
                log.warn("第 {} 行分配失败，回退到单行独立拆分", rowIndex + 2);
                rowAllocations.set(rowIndex, null);
            }
        }

        return new CardAllocationContext(rowAllocations);
    }

    /**
     * 从库存中分配 target 数量，返回 (卡号, 数量) 列表。
     * 只从现有卡号库存扣减，不重新拆分，保证每个卡号「收货=发货」且总数量一致；
     * 单张提单内同一卡号只出现一次，数量互不重复。
     */
    private List<CardQty> allocateFromInventory(int target, Map<String, Integer> inventory,
                                                int min, int max, int maxSubCount) {
        if (target <= 0 || inventory.isEmpty()) return Collections.emptyList();

        int totalAvail = inventory.values().stream().mapToInt(Integer::intValue).sum();
        if (totalAvail < target) return null;

        // 按数量从大到小排序，优先用大卡
        List<Map.Entry<String, Integer>> cards = new ArrayList<>(inventory.entrySet());
        cards.sort((a, b) -> Integer.compare(b.getValue(), a.getValue()));

        List<CardQty> result = new ArrayList<>();
        Set<Integer> usedAmounts = new HashSet<>();
        int sum = 0;

        for (Map.Entry<String, Integer> e : cards) {
            if (sum == target) break;
            String cardNo = e.getKey();
            int qty = e.getValue();
            int need = target - sum;

            if (need <= 0) break;

            int takeAmt;
            if (qty <= need) {
                takeAmt = qty;
            } else {
                takeAmt = need;
                while (usedAmounts.contains(takeAmt) && takeAmt < qty) takeAmt++;
                if (usedAmounts.contains(takeAmt)) {
                    takeAmt = need;
                    while (takeAmt > 0 && usedAmounts.contains(takeAmt)) takeAmt--;
                    if (takeAmt <= 0) continue;
                }
            }

            result.add(new CardQty(cardNo, takeAmt));
            usedAmounts.add(takeAmt);
            sum += takeAmt;
        }

        if (sum != target) {
            if (sum < target) return null;
            int diff = sum - target;
            for (int i = result.size() - 1; i >= 0; i--) {
                CardQty cq = result.get(i);
                if (cq.qty <= diff) continue;
                int newQty = cq.qty - diff;
                usedAmounts.remove(cq.qty);
                if (!usedAmounts.contains(newQty)) {
                    result.set(i, new CardQty(cq.cardNo, newQty));
                    usedAmounts.add(newQty);
                    sum = target;
                    break;
                }
                usedAmounts.add(cq.qty);
            }
            if (sum != target) return null;
        }

        return result.isEmpty() ? null : result;
    }

    private static class RowRef {
        final int rowIndex;
        final int qty;
        final Map<String, String> rowData;

        RowRef(int rowIndex, int qty, Map<String, String> rowData) {
            this.rowIndex = rowIndex;
            this.qty = qty;
            this.rowData = rowData;
        }
    }

    private static class CardQty {
        final String cardNo;
        final int qty;

        CardQty(String cardNo, int qty) {
            this.cardNo = cardNo;
            this.qty = qty;
        }
    }

    private static class CardAllocationContext {
        final List<List<CardQty>> rowAllocations;

        CardAllocationContext(List<List<CardQty>> rowAllocations) {
            this.rowAllocations = rowAllocations;
        }

        List<CardQty> getAllocation(int rowIndex) {
            if (rowIndex < 0 || rowIndex >= rowAllocations.size()) return null;
            return rowAllocations.get(rowIndex);
        }
    }

    /**
     * 提单编号：供方/发货单位前4字拼音-需方/提货单位前4字拼音-签发时间-001
     */
    private String generateBillNumber(Map<String, String> rowData, Map<String, Integer> seqMap) {
        String gongFang = firstFourChars(emptyToBlank(getColumnValue(rowData, "供方/发货单位", "供方／发货单位")));
        String xuFang = firstFourChars(emptyToBlank(getColumnValue(rowData, "需方/提货单位", "需方／提货单位")));
        String dateStr = emptyToBlank(getColumnValue(rowData, "签发时间"));
        String gongInitials = toPinyinInitials(gongFang).toUpperCase();
        String xuInitials = toPinyinInitials(xuFang).toUpperCase();
        if (gongInitials.isEmpty()) gongInitials = "X";
        if (xuInitials.isEmpty()) xuInitials = "X";
        String yyyyMMdd = formatDate(dateStr);
        String groupKey = gongFang + "|" + xuFang + "|" + yyyyMMdd;
        int seq = seqMap.merge(groupKey, 1, Integer::sum);
        return gongInitials + "-" + xuInitials + "-" + yyyyMMdd + "-" + String.format("%03d", seq);
    }

    /** 按多个可能的列名取值，兼容全角/半角斜杠等 */
    private String getColumnValue(Map<String, String> rowData, String... possibleKeys) {
        for (String key : possibleKeys) {
            String v = rowData.get(key);
            if (v != null && !v.trim().isEmpty()) return v;
        }
        for (Map.Entry<String, String> e : rowData.entrySet()) {
            String k = e.getKey();
            if (k == null) continue;
            for (String pk : possibleKeys) {
                if (normalizeKey(k).equals(normalizeKey(pk))) return e.getValue();
            }
        }
        return rowData.get(possibleKeys[0]);
    }

    private String normalizeKey(String key) {
        if (key == null) return "";
        return key.trim().replace("／", "/").replace("\u00A0", " ");
    }

    private String firstFourChars(String s) {
        if (s == null) return "";
        s = s.trim();
        int len = 0;
        for (int i = 0; i < s.length() && len < 4; i++) {
            len++;
        }
        return s.substring(0, Math.min(len, s.length()));
    }

    private String emptyToBlank(String s) {
        return s == null ? "" : s.trim();
    }

    private String formatDate(String dateStr) {
        if (dateStr == null || dateStr.trim().isEmpty()) {
            return LocalDate.now().format(DateTimeFormatter.BASIC_ISO_DATE);
        }
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

    /**
     * 单行独立拆分（回退逻辑）：将数量拆分为若干份，每份在 [splitMin, splitMax] 内互不重复
     */
    private List<Double> splitQuantity(double total) {
        if (total <= 0) return new ArrayList<>();
        int totalInt = (int) Math.round(total);
        int min = config.getSplitMin();
        int max = config.getSplitMax();
        int maxSubCount = config.getSplitMaxSubCount();
        if (totalInt < min) {
            return Collections.singletonList((double) totalInt);
        }
        try {
            List<Integer> parts = OrderSplitUtil.split(totalInt, min, max, maxSubCount);
            List<Double> result = new ArrayList<>(parts.size());
            for (Integer p : parts) {
                result.add(p.doubleValue());
            }
            return result;
        } catch (IllegalArgumentException e) {
            return Collections.singletonList((double) totalInt);
        }
    }

    private Workbook openWorkbook(InputStream is, Path file) throws IOException {
        PushbackInputStream pis = new PushbackInputStream(is, 8);
        byte[] header = new byte[4];
        int n = pis.read(header);
        if (n > 0) pis.unread(header, 0, n);
        boolean isXlsx = n >= 2 && header[0] == 0x50 && header[1] == 0x4B;
        return isXlsx ? new XSSFWorkbook(pis) : new HSSFWorkbook(pis);
    }

    private List<Map<String, String>> readDataSheet(Path dataFile) throws IOException {
        List<Map<String, String>> rows = new ArrayList<>();
        try (InputStream is = Files.newInputStream(dataFile);
             Workbook workbook = openWorkbook(is, dataFile)) {
            Sheet sheet = workbook.getSheetAt(0);
            if (sheet == null || sheet.getPhysicalNumberOfRows() < 2) return rows;
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
                    if (header == null || header.trim().isEmpty()) header = "col" + c;
                    Cell cell = row.getCell(c);
                    rowData.put(header.trim(), cell != null ? getCellStringValue(cell, header) : "");
                }
                rows.add(rowData);
            }
        }
        return rows;
    }

    private String getCellStringValue(Cell cell, String columnHeader) {
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getLocalDateTimeCellValue().toLocalDate().format(DateTimeFormatter.ofPattern("yyyy/M/d"));
                }
                return formatNumeric(cell.getNumericCellValue(), -1);
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                try {
                    return formatNumeric(cell.getNumericCellValue(), -1);
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

    private static String formatNumeric(double num, int decimals) {
        if (Double.isNaN(num) || Double.isInfinite(num)) return String.valueOf(num);
        if (decimals == 0 || (decimals < 0 && num == Math.floor(num) && Math.abs(num) < 1e15)) {
            return String.valueOf((long) num);
        }
        int scale = decimals >= 0 ? decimals : 6;
        BigDecimal bd = BigDecimal.valueOf(num).setScale(scale, RoundingMode.HALF_UP);
        return bd.stripTrailingZeros().toPlainString();
    }

    private void fillAndSave(Path templateFile, Map<String, String> rowData, Path outputFile,
                             CardAllocationContext ctx, int rowIndex) throws IOException {
        try (InputStream is = Files.newInputStream(templateFile)) {
            Workbook workbook = openWorkbook(is, templateFile);

            replacePlaceholders(workbook, rowData);
            fillDetailTable(workbook, rowData, ctx, rowIndex);

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

    private Map<String, String> buildReplacementMap(Map<String, String> rowData) {
        Map<String, String> map = new HashMap<>(rowData);
        for (Map.Entry<String, String> e : new HashMap<>(rowData).entrySet()) {
            String normKey = normalizeKey(e.getKey());
            if (!normKey.equals(e.getKey())) map.put(normKey, e.getValue());
        }
        String gongFang = emptyToBlank(getColumnValue(rowData, "供方/发货单位", "供方／发货单位"));
        String xuFang = emptyToBlank(getColumnValue(rowData, "需方/提货单位", "需方／提货单位"));
        String issueDate = emptyToBlank(getColumnValue(rowData, "签发时间"));
        map.put("供方/发货单位", gongFang);
        map.put("供方／发货单位", gongFang);
        map.put("需方/提货单位", xuFang);
        map.put("需方／提货单位", xuFang);
        map.put("签发时间", issueDate);
        return map;
    }

    private void replacePlaceholders(Workbook workbook, Map<String, String> rowData) {
        Map<String, String> replaceMap = buildReplacementMap(rowData);
        String prefix = config.getPlaceholderPrefix();
        String suffix = config.getPlaceholderSuffix();
        Pattern pattern = Pattern.compile(Pattern.quote(prefix) + "([^" + Pattern.quote(suffix) + "]+)" + Pattern.quote(suffix));

        for (int s = 0; s < workbook.getNumberOfSheets(); s++) {
            Sheet sheet = workbook.getSheetAt(s);
            for (Row row : sheet) {
                if (row == null) continue;
                for (Cell cell : row) {
                    if (cell == null) continue;
                    if (cell.getCellType() == CellType.STRING) {
                        String value = cell.getStringCellValue();
                        String replaced = replaceInString(value, replaceMap, pattern, prefix, suffix);
                        if (!value.equals(replaced)) {
                            cell.setCellValue(replaced);
                        }
                    } else if (cell.getCellType() == CellType.FORMULA) {
                        try {
                            String formula = cell.getCellFormula();
                            if (pattern.matcher(formula).find()) {
                                String replaced = replaceInString(formula, replaceMap, pattern, prefix, suffix);
                                if (!formula.equals(replaced)) {
                                    cell.setCellFormula(replaced);
                                }
                            }
                        } catch (Exception ignored) {}
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
            String replacement = replaceMap.getOrDefault(key, replaceMap.getOrDefault(normalizeKey(key), ""));
            matcher.appendReplacement(sb, Matcher.quoteReplacement(replacement));
        }
        matcher.appendTail(sb);
        return sb.toString();
    }

    /**
     * 查找表头行并填充数据行。优先使用预计算的卡号分配（保证跨单卡号对应），否则回退到单行独立拆分。
     */
    private void fillDetailTable(Workbook workbook, Map<String, String> rowData,
                                 CardAllocationContext ctx, int rowIndex) {
        Sheet sheet = workbook.getSheetAt(0);
        int headerRowIdx = findTableHeaderRow(sheet);
        if (headerRowIdx < 0) {
            log.warn("未找到表头行（产品名、规格型号、卡号、数量（吨）、件数、备注），跳过表格填充");
            return;
        }

        int[] colIndices = findTableColumnIndices(sheet.getRow(headerRowIdx));
        if (colIndices == null) return;

        String productName = emptyToBlank(getColumnValue(rowData, "产品名"));
        String specModel = emptyToBlank(getColumnValue(rowData, "规格型号"));
        String xuFang = emptyToBlank(getColumnValue(rowData, "需方/提货单位", "需方／提货单位"));

        List<CardQty> allocation = ctx != null ? ctx.getAllocation(rowIndex) : null;
        if (allocation != null && !allocation.isEmpty()) {
            // 为了可读性，按卡号尾号从小到大排序，例如 52518-1, 52518-2, ...
            allocation.sort((a, b) -> {
                String ca = a.cardNo;
                String cb = b.cardNo;
                int ia = ca.lastIndexOf('-');
                int ib = cb.lastIndexOf('-');
                String sa = ia >= 0 ? ca.substring(ia + 1) : ca;
                String sb = ib >= 0 ? cb.substring(ib + 1) : cb;
                try {
                    int na = Integer.parseInt(sa);
                    int nb = Integer.parseInt(sb);
                    return Integer.compare(na, nb);
                } catch (NumberFormatException e) {
                    return ca.compareTo(cb);
                }
            });
        }
        List<Double> qtysForFallback = null;

        int firstDataRowIdx = headerRowIdx + 1;
        int templateLastRowIdx = firstDataRowIdx + 28;
        int dataRowCount;

        if (allocation == null || allocation.isEmpty()) {
            qtysForFallback = splitQuantity(parseQuantity(rowData));
            dataRowCount = qtysForFallback.isEmpty() ? 0 : qtysForFallback.size();
        } else {
            dataRowCount = allocation.size();
        }

        if (dataRowCount == 0) {
            writeTableRow(ensureRow(sheet, firstDataRowIdx), colIndices, productName, specModel,
                    specModel + "-1", 0, 0, "请过户至" + xuFang + specModel + "-1");
            dataRowCount = 1;
        } else {
            Row styleSourceRow = sheet.getRow(firstDataRowIdx);
            for (int i = 0; i < dataRowCount; i++) {
                double qty;
                int pieceCount;
                String cardNo;
                String remark;
                if (allocation != null && !allocation.isEmpty()) {
                    CardQty cq = allocation.get(i);
                    qty = cq.qty;
                    pieceCount = cq.qty * PIECES_MULTIPLIER;
                    cardNo = cq.cardNo;
                    remark = "请过户至" + xuFang + cq.cardNo;
                } else {
                    // 回退路径：使用预先计算好的 qtysForFallback，避免每次循环重新随机拆分导致长度不一致
                    qty = qtysForFallback.get(i);
                    pieceCount = (int) Math.round(qty * PIECES_MULTIPLIER);
                    cardNo = specModel + "-" + (i + 1);
                    remark = "请过户至" + xuFang + cardNo;
                }
                int targetRowIdx = firstDataRowIdx + i;
                Row row;
                if (targetRowIdx <= templateLastRowIdx) {
                    row = ensureRow(sheet, targetRowIdx);
                } else {
                    row = insertRowWithStyle(sheet, targetRowIdx, styleSourceRow);
                }
                writeTableRow(row, colIndices, productName, specModel, cardNo, qty, pieceCount, remark);
            }
        }
        int lastDataRowIdx = firstDataRowIdx + dataRowCount - 1;
        updateSumFormulasToLastRow(sheet, lastDataRowIdx);
    }

    private int findTableHeaderRow(Sheet sheet) {
        for (int r = 0; r <= Math.min(sheet.getLastRowNum(), 30); r++) {
            Row row = sheet.getRow(r);
            if (row == null) continue;
            for (Cell cell : row) {
                if (cell == null) continue;
                String val = getCellStringValue(cell, null);
                if (val != null && val.contains("产品名")) {
                    return r;
                }
            }
        }
        return -1;
    }

    private int[] findTableColumnIndices(Row headerRow) {
        if (headerRow == null) return null;
        Map<String, Integer> headerToCol = new HashMap<>();
        for (Cell cell : headerRow) {
            String val = getCellStringValue(cell, null);
            if (val == null) continue;
            val = val.trim();
            for (String h : TABLE_HEADERS) {
                if (val.contains(h) || h.contains(val)) {
                    headerToCol.putIfAbsent(h, cell.getColumnIndex());
                    break;
                }
            }
        }
        int[] indices = new int[TABLE_HEADERS.length];
        for (int i = 0; i < TABLE_HEADERS.length; i++) {
            Integer col = headerToCol.get(TABLE_HEADERS[i]);
            if (col == null) return null;
            indices[i] = col;
        }
        return indices;
    }

    private double parseQuantity(Map<String, String> rowData) {
        String v = getColumnValue(rowData, "数量", "数量（吨）");
        if (v == null || v.trim().isEmpty()) return 0;
        try {
            return Double.parseDouble(v.trim().replace(",", ""));
        } catch (NumberFormatException e) {
            return 0;
        }
    }

    private void writeTableRow(Row row, int[] colIndices, String productName, String specModel,
                              String cardNo, double qty, int pieceCount, String remark) {
        if (row == null) return;
        setCellValue(getCell(row, colIndices[0]), productName);
        setCellValue(getCell(row, colIndices[1]), specModel);
        setCellValue(getCell(row, colIndices[2]), cardNo);
        setCellValue(getCell(row, colIndices[3]), qty);
        setCellValue(getCell(row, colIndices[4]), pieceCount);
        setCellValue(getCell(row, colIndices[5]), remark);
    }

    private Row ensureRow(Sheet sheet, int rowIdx) {
        Row row = sheet.getRow(rowIdx);
        return row != null ? row : sheet.createRow(rowIdx);
    }

    private Row insertRowWithStyle(Sheet sheet, int insertAt, Row styleSourceRow) {
        sheet.shiftRows(insertAt, sheet.getLastRowNum(), 1, true, false);
        Row newRow = sheet.createRow(insertAt);
        if (styleSourceRow != null) {
            int srcRowNum = styleSourceRow.getRowNum();
            newRow.setHeight(styleSourceRow.getHeight());
            short lastCol = styleSourceRow.getLastCellNum();
            for (short c = 0; c < lastCol; c++) {
                Cell srcCell = styleSourceRow.getCell(c);
                if (srcCell != null && srcCell.getCellStyle() != null) {
                    Cell newCell = newRow.createCell(c);
                    newCell.setCellStyle(srcCell.getCellStyle());
                }
            }
            copyMergedRegionsToRow(sheet, srcRowNum, insertAt);
        }
        return newRow;
    }

    private void copyMergedRegionsToRow(Sheet sheet, int sourceRow, int newRow) {
        List<CellRangeAddress> toAdd = new ArrayList<>();
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress r = sheet.getMergedRegion(i);
            if (r.getFirstRow() <= sourceRow && sourceRow <= r.getLastRow()) {
                toAdd.add(new CellRangeAddress(newRow, newRow, r.getFirstColumn(), r.getLastColumn()));
            }
        }
        for (CellRangeAddress r : toAdd) {
            sheet.addMergedRegion(r);
        }
    }

    private void updateSumFormulasToLastRow(Sheet sheet, int lastDataRowIdx) {
        int lastRowExcel = lastDataRowIdx + 1;
        Pattern rangePattern = Pattern.compile("(SUM\\()([^)]+)(\\))");
        for (Row row : sheet) {
            if (row == null) continue;
            for (Cell cell : row) {
                if (cell == null || cell.getCellType() != CellType.FORMULA) continue;
                try {
                    String formula = cell.getCellFormula();
                    if (!formula.toUpperCase().contains("SUM(")) continue;
                    Matcher m = rangePattern.matcher(formula);
                    StringBuffer sb = new StringBuffer();
                    while (m.find()) {
                        String range = m.group(2);
                        String updated = range.replaceAll(":([A-Z$]+)\\d+$", ":$1" + lastRowExcel);
                        m.appendReplacement(sb, Matcher.quoteReplacement(m.group(1) + updated + m.group(3)));
                    }
                    m.appendTail(sb);
                    String newFormula = sb.toString();
                    if (!newFormula.equals(formula)) {
                        cell.setCellFormula(newFormula);
                    }
                } catch (Exception e) {
                    log.debug("跳过公式更新: {}", e.getMessage());
                }
            }
        }
    }

    private Cell getCell(Row row, int colIdx) {
        Cell cell = row.getCell(colIdx);
        if (cell == null) cell = row.createCell(colIdx);
        return cell;
    }

    private void setCellValue(Cell cell, String value) {
        if (cell == null) return;
        if (value == null) value = "";
        value = value.trim();
        try {
            cell.setCellValue(Double.parseDouble(value.replace(",", "")));
        } catch (NumberFormatException e) {
            cell.setCellValue(value);
        }
    }

    private void setCellValue(Cell cell, double value) {
        if (cell == null) return;
        cell.setCellValue(value);
    }

    private void setCellValue(Cell cell, int value) {
        if (cell == null) return;
        cell.setCellValue(value);
    }

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

    private String ensureUniqueFileName(String fileName, Set<String> usedFileNames, int rowIndex) {
        if (!usedFileNames.contains(fileName)) return fileName;
        int lastDot = fileName.lastIndexOf('.');
        String base = lastDot > 0 ? fileName.substring(0, lastDot) : fileName;
        String ext = lastDot > 0 ? fileName.substring(lastDot) : "";
        return base + "_" + rowIndex + ext;
    }
}
