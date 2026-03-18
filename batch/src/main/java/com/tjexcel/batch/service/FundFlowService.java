package com.tjexcel.batch.service;

import com.tjexcel.batch.config.FundFlowConfig;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.*;

/**
 * 资金流生成（无模板）：每个数据 sheet 输出 1 个资金流表（一个输出文件内多 Sheet）
 * 方向：需方(付款方) -> 供方(收款方)
 * 金额：默认取配置 amountColumn（一般为 价税合计金额）
 */
@Service
public class FundFlowService {

    private static final Logger log = LoggerFactory.getLogger(FundFlowService.class);

    private final FundFlowConfig config;

    public FundFlowService(FundFlowConfig config) {
        this.config = config;
    }

    public int generate() throws IOException {
        Path dataFile = Paths.get(config.getDataPath()).toAbsolutePath();
        Path outputDir = Paths.get(config.getOutputDir()).toAbsolutePath();
        Files.createDirectories(outputDir);

        if (!Files.exists(dataFile)) {
            throw new FileNotFoundException("数据表不存在: " + dataFile);
        }

        int sheetSuccessCount = 0;
        try (InputStream is = Files.newInputStream(dataFile);
             Workbook dataWb = openWorkbook(is, dataFile)) {

            for (int sheetIndex = 0; sheetIndex < dataWb.getNumberOfSheets(); sheetIndex++) {
                Sheet sheet = dataWb.getSheetAt(sheetIndex);
                if (sheet == null) continue;

                List<RowItem> items = readFundFlowItems(sheet);
                if (items.isEmpty()) {
                    log.warn("Sheet [{}] 无有效资金流数据，跳过", sheet.getSheetName());
                    continue;
                }

                String outName = buildOutputFileName(sheet.getSheetName());
                Path outFile = outputDir.resolve(outName);

                try (Workbook outWb = new XSSFWorkbook()) {
                    String outSheetName = sanitizeSheetName(sheet.getSheetName());
                    if (outSheetName.isEmpty()) outSheetName = "资金流";
                    Sheet outSheet = outWb.createSheet(outSheetName);
                    writeFundFlowSheet(outWb, outSheet, items);

                    if (outWb instanceof XSSFWorkbook) {
                        ((XSSFWorkbook) outWb).setForceFormulaRecalculation(true);
                    } else if (outWb instanceof HSSFWorkbook) {
                        ((HSSFWorkbook) outWb).setForceFormulaRecalculation(true);
                    }

                    try (OutputStream os = Files.newOutputStream(outFile)) {
                        outWb.write(os);
                    }
                }

                log.info("Sheet [{}] 已生成资金流文件: {} ({} 条明细)", sheet.getSheetName(), outFile.getFileName(), items.size());
                sheetSuccessCount++;
            }
        }

        log.info("资金流生成完成，共生成 {} 个Sheet，输出目录: {}", sheetSuccessCount, outputDir);
        return sheetSuccessCount;
    }

    private String buildOutputFileName(String sheetName) {
        String safe = sanitizeFileName(sheetName);
        if (safe.isEmpty()) safe = "sheet";
        return "资金流-" + safe + ".xlsx";
    }

    /**
     * 横向资金流：一层一列，付款方在左；同层多公司同一列，纵向对齐（客户与公司、公司与上游对齐）。
     * 表头每列只写一次；列内按笔数显示（对方名+金额），几笔几行。
     */
    private void writeFundFlowSheet(Workbook wb, Sheet out, List<RowItem> items) {
        DataFormat df = wb.createDataFormat();
        CellStyle nameStyle = wb.createCellStyle();
        nameStyle.setAlignment(HorizontalAlignment.LEFT);
        nameStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        CellStyle amountStyle = wb.createCellStyle();
        amountStyle.setDataFormat(df.getFormat("#,##0.00"));
        amountStyle.setAlignment(HorizontalAlignment.LEFT);
        amountStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        // 1. 建图：保持输入顺序
        Map<String, List<RowItem>> outEdges = new LinkedHashMap<>();
        LinkedHashSet<String> payerOrder = new LinkedHashSet<>();
        Set<String> incoming = new HashSet<>();
        for (RowItem it : items) {
            payerOrder.add(it.payer);
            incoming.add(it.payee);
            outEdges.computeIfAbsent(it.payer, k -> new ArrayList<>()).add(it);
        }

        // 2. 根节点：入度为 0 的付款方（保持出现顺序）
        List<String> roots = new ArrayList<>();
        for (String payer : payerOrder) {
            if (!incoming.contains(payer)) roots.add(payer);
        }
        if (roots.isEmpty() && !payerOrder.isEmpty()) {
            roots.add(payerOrder.iterator().next());
        }

        // 3. 构建树并写表：父节点显示一次，子节点多笔在其下方展开
        int[] maxDepth = new int[] {0};
        int rowPairCursor = 0;
        for (String root : roots) {
            FlowNode node = buildFlowNode(root, null, outEdges, new LinkedHashSet<>());
            rowPairCursor = writeFlowNode(out, node, nameStyle, amountStyle, 0, rowPairCursor, maxDepth);
        }

        int numCols = maxDepth[0] + 1;
        for (int c = 0; c < numCols; c++) {
            out.autoSizeColumn(c);
            int w = out.getColumnWidth(c);
            int doubled = (int) Math.min((long) w * 2L, 255L * 256L);
            out.setColumnWidth(c, Math.max(20 * 256, doubled));
        }
    }

    private static class FlowNode {
        final String company;
        final BigDecimal amount;
        final List<FlowNode> children = new ArrayList<>();

        FlowNode(String company, BigDecimal amount) {
            this.company = company;
            this.amount = amount;
        }
    }

    private FlowNode buildFlowNode(String company,
                                   BigDecimal amount,
                                   Map<String, List<RowItem>> outEdges,
                                   LinkedHashSet<String> path) {
        FlowNode node = new FlowNode(company, amount);
        if (path.contains(company)) {
            return node;
        }
        LinkedHashSet<String> nextPath = new LinkedHashSet<>(path);
        nextPath.add(company);
        List<RowItem> outs = outEdges.getOrDefault(company, Collections.emptyList());
        for (RowItem it : outs) {
            node.children.add(buildFlowNode(it.payee, it.amount, outEdges, nextPath));
        }
        return node;
    }

    private int writeFlowNode(Sheet out,
                              FlowNode node,
                              CellStyle nameStyle,
                              CellStyle amountStyle,
                              int depth,
                              int startRowPair,
                              int[] maxDepth) {
        maxDepth[0] = Math.max(maxDepth[0], depth);
        writeCell(out, depth, startRowPair, node.company, node.amount, nameStyle, amountStyle);
        int row = startRowPair;
        if (node.children.isEmpty()) {
            return startRowPair + 1;
        }
        for (FlowNode ch : node.children) {
            row = writeFlowNode(out, ch, nameStyle, amountStyle, depth + 1, row, maxDepth);
        }
        return row;
    }

    private void writeCell(Sheet out,
                           int col,
                           int rowPair,
                           String name,
                           BigDecimal amount,
                           CellStyle nameStyle,
                           CellStyle amountStyle) {
        int rowName = rowPair * 2;
        int rowAmt = rowName + 1;
        Row nameRow = out.getRow(rowName);
        if (nameRow == null) nameRow = out.createRow(rowName);
        Row amountRow = out.getRow(rowAmt);
        if (amountRow == null) amountRow = out.createRow(rowAmt);

        Cell cn = nameRow.getCell(col);
        if (cn == null) cn = nameRow.createCell(col);
        cn.setCellStyle(nameStyle);
        cn.setCellValue(name == null ? "" : name);

        if (amount != null) {
            Cell ca = amountRow.getCell(col);
            if (ca == null) ca = amountRow.createCell(col);
            ca.setCellStyle(amountStyle);
            ca.setCellValue(amount.doubleValue());
        }
    }

    private List<RowItem> readFundFlowItems(Sheet sheet) {
        List<RowItem> items = new ArrayList<>();
        if (sheet.getPhysicalNumberOfRows() < 2) return items;

        Row headerRow = sheet.getRow(0);
        if (headerRow == null) return items;

        List<String> headers = new ArrayList<>();
        for (Cell cell : headerRow) {
            headers.add(getCellStringValue(cell, null));
        }

        int payerIdx = findHeaderIndex(headers, Arrays.asList("需方", "买方", "需方/提货单位", "需方／提货单位"));
        int payeeIdx = findHeaderIndex(headers, Arrays.asList("供方", "卖方", "供方/发货单位", "供方／发货单位"));
        int amountIdx = findHeaderIndex(headers, Collections.singletonList(config.getAmountColumn()));

        // 兜底：找不到金额列时，尝试常见列名
        if (amountIdx < 0) {
            amountIdx = findHeaderIndex(headers, Arrays.asList("价税合计金额", "不含税金额", "结算金额"));
        }

        for (int r = 1; r <= sheet.getLastRowNum(); r++) {
            Row row = sheet.getRow(r);
            if (row == null) continue;

            String payer = getCellStringValue(row.getCell(payerIdx), "需方").trim();
            String payee = getCellStringValue(row.getCell(payeeIdx), "供方").trim();
            String amountStr = getCellStringValue(row.getCell(amountIdx), config.getAmountColumn()).trim();

            if (payer.isEmpty() || payee.isEmpty()) continue;
            BigDecimal amount = parseAmount(amountStr);
            if (amount == null || amount.compareTo(BigDecimal.ZERO) <= 0) continue;

            items.add(new RowItem(payer, payee, amount, r + 1));
        }
        return items;
    }

    private static int findHeaderIndex(List<String> headers, List<String> candidates) {
        if (headers == null) return -1;
        for (int i = 0; i < headers.size(); i++) {
            String h = headers.get(i) == null ? "" : headers.get(i).trim();
            for (String c : candidates) {
                if (c != null && !c.trim().isEmpty() && h.equals(c.trim())) {
                    return i;
                }
            }
        }
        return -1;
    }

    private static BigDecimal parseAmount(String s) {
        if (s == null) return null;
        String t = s.trim().replace(",", "");
        if (t.isEmpty()) return null;
        try {
            return new BigDecimal(t).setScale(2, RoundingMode.HALF_UP);
        } catch (Exception ignored) {
            return null;
        }
    }

    private static Workbook openWorkbook(InputStream is, Path file) throws IOException {
        String name = file.getFileName().toString().toLowerCase();
        if (name.endsWith(".xlsx")) return new XSSFWorkbook(is);
        return new HSSFWorkbook(is);
    }

    private static String sanitizeFileName(String name) {
        if (name == null) return "";
        return name.replaceAll("[\\\\/:*?\"<>|]", "_").trim();
    }

    private static String sanitizeSheetName(String name) {
        String safe = sanitizeFileName(name);
        if (safe.length() > 31) safe = safe.substring(0, 31);
        return safe;
    }

    private static final DateTimeFormatter DATE_DISPLAY_FMT = DateTimeFormatter.ofPattern("yyyy/M/d");

    private String getCellStringValue(Cell cell, String columnHeader) {
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    try {
                        LocalDate d = cell.getLocalDateTimeCellValue().toLocalDate();
                        return d.format(DATE_DISPLAY_FMT);
                    } catch (Exception ignored) {
                        return "";
                    }
                }
                return formatNumeric(cell.getNumericCellValue(), isAmountColumn(columnHeader) ? 2 : -1);
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                try {
                    if (DateUtil.isCellDateFormatted(cell)) {
                        LocalDate d = cell.getLocalDateTimeCellValue().toLocalDate();
                        return d.format(DATE_DISPLAY_FMT);
                    }
                    return formatNumeric(cell.getNumericCellValue(), isAmountColumn(columnHeader) ? 2 : -1);
                } catch (Exception e) {
                    try {
                        return cell.getStringCellValue();
                    } catch (Exception ignored) {
                        return cell.getCellFormula();
                    }
                }
            default:
                return "";
        }
    }

    private boolean isAmountColumn(String header) {
        if (header == null) return false;
        String h = header.trim();
        return h.equals("价税合计金额") || h.equals("不含税金额") || h.equals("增值税额") || h.equals(config.getAmountColumn());
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

    private static class RowItem {
        final String payer;
        final String payee;
        final BigDecimal amount;
        final int sourceRowNo;

        RowItem(String payer, String payee, BigDecimal amount, int sourceRowNo) {
            this.payer = payer;
            this.payee = payee;
            this.amount = amount;
            this.sourceRowNo = sourceRowNo;
        }
    }

    // PairKey/Summary 已不再使用（旧版“明细+汇总”布局）
}


