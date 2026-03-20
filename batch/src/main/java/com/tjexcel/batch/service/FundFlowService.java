package com.tjexcel.batch.service;

import com.tjexcel.batch.config.FundFlowConfig;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;

import java.io.*;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.nio.file.*;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.*;

/**
 * 资金流生成服务。
 *
 * 方向：需方（付款方）-> 供方（收款方），金额取"价税合计金额"列。
 * 按 README 规则 1-10 生成资金流表格。
 */
@Service
public class FundFlowService {

    private static final Logger log = LoggerFactory.getLogger(FundFlowService.class);

    private final FundFlowConfig config;

    public FundFlowService(FundFlowConfig config) {
        this.config = config;
    }

    // =========================================================================
    // 入口
    // =========================================================================

    public int generate() throws IOException {
        Path dataFile = Paths.get(config.getDataPath()).toAbsolutePath();
        Path outputDir = Paths.get(config.getOutputDir()).toAbsolutePath();
        Files.createDirectories(outputDir);

        if (!Files.exists(dataFile)) {
            throw new FileNotFoundException("数据表不存在: " + dataFile);
        }

        int count = 0;
        try (InputStream is = Files.newInputStream(dataFile);
             Workbook dataWb = openWorkbook(is, dataFile)) {

            for (int si = 0; si < dataWb.getNumberOfSheets(); si++) {
                Sheet sheet = dataWb.getSheetAt(si);
                if (sheet == null) continue;

                List<Edge> edges = readEdges(sheet);
                if (edges.isEmpty()) {
                    log.warn("Sheet [{}] 无有效资金流数据，跳过", sheet.getSheetName());
                    continue;
                }

                String outName = "资金流_" + sanitizeFileName(sheet.getSheetName()) + ".xlsx";
                Path outFile = outputDir.resolve(outName);

                try (XSSFWorkbook outWb = new XSSFWorkbook()) {
                    String sheetName = sanitizeSheetName(sheet.getSheetName());
                    if (sheetName.isEmpty()) sheetName = "资金流";
                    Sheet outSheet = outWb.createSheet(sheetName);
                    writeSheet(outWb, outSheet, edges);
                    outWb.setForceFormulaRecalculation(true);
                    try (OutputStream os = Files.newOutputStream(outFile)) {
                        outWb.write(os);
                    }
                }

                log.info("Sheet [{}] -> {} ({} 条明细)", sheet.getSheetName(), outFile.getFileName(), edges.size());
                count++;
            }
        }

        log.info("资金流生成完成，共 {} 个文件，输出目录: {}", count, outputDir);
        return count;
    }

    // =========================================================================
    // 数据模型
    // =========================================================================

    /** 一条有向交易记录（有向边）。rowNo 为数据表原始行号（1-based），用于排序和唯一性。*/
    private static class Edge {
        final String payer;       // 需方（付款方）
        final String payee;       // 供方（收款方）
        final BigDecimal amount;
        final int rowNo;

        Edge(String payer, String payee, BigDecimal amount, int rowNo) {
            this.payer = payer;
            this.payee = payee;
            this.amount = amount;
            this.rowNo = rowNo;
        }
    }

    private static class Chain {
        final List<Edge> edges = new ArrayList<>();
    }

    private static class ChainWriteResult {
        final int cols;  // 使用的列数
        final int rows;  // 使用的行数
        
        ChainWriteResult(int cols, int rows) {
            this.cols = cols;
            this.rows = rows;
        }
    }

    // =========================================================================
    // Excel 写入
    // =========================================================================

    private void writeSheet(Workbook wb, Sheet out, List<Edge> allEdges) {
        CellStyle nameStyle = createNameStyle(wb);
        CellStyle amtStyle  = createAmountStyle(wb);

        List<Chain> chains = buildChains(allEdges);
        int rowCursor = 0;
        int maxCol = 0;

        for (Chain chain : chains) {
            ChainWriteResult result = writeLinearChain(out, chain, rowCursor, nameStyle, amtStyle);
            if (result.cols > 0) {
                maxCol = Math.max(maxCol, result.cols - 1);
                rowCursor += result.rows; // 名称行 + 金额行（可能多行）
                rowCursor++;              // 链间空行
            }
        }

        for (int c = 0; c <= maxCol; c++) {
            out.autoSizeColumn(c);
            int w = out.getColumnWidth(c);
            int adjusted = (int) Math.min((long) w * 3L / 2L, 255L * 256L);
            out.setColumnWidth(c, Math.max(20 * 256, adjusted));
        }
    }

    private List<Chain> buildChains(List<Edge> allEdges) {
        List<Edge> edges = new ArrayList<>(allEdges);
        edges.sort(Comparator.comparingInt(e -> e.rowNo));

        Map<String, List<Edge>> byPayee = new LinkedHashMap<>();
        for (Edge e : edges) {
            byPayee.computeIfAbsent(e.payee, k -> new ArrayList<>()).add(e);
        }

        Set<Integer> usedRows = new LinkedHashSet<>();
        Deque<String> queue = new ArrayDeque<>();
        if (!edges.isEmpty()) {
            queue.add(edges.get(0).payee);
        }

        List<Chain> chains = new ArrayList<>();
        while (true) {
            while (!queue.isEmpty()) {
                String startPayee = queue.pollFirst();
                Chain chain = buildChainFromPayee(startPayee, byPayee, usedRows, queue);
                if (!chain.edges.isEmpty()) {
                    chains.add(chain);
                }
            }
            Edge next = firstUnusedEdge(edges, usedRows);
            if (next == null) break;
            queue.add(next.payee);
        }

        return chains;
    }

    private Edge firstUnusedEdge(List<Edge> edges, Set<Integer> usedRows) {
        for (Edge e : edges) {
            if (!usedRows.contains(e.rowNo)) return e;
        }
        return null;
    }

    private Chain buildChainFromPayee(String startPayee,
                                      Map<String, List<Edge>> byPayee,
                                      Set<Integer> usedRows,
                                      Deque<String> queue) {
        Chain chain = new Chain();
        String current = startPayee;
        LinkedHashSet<String> path = new LinkedHashSet<>();

        while (true) {
            if (!path.add(current)) break; // 避免闭环

            List<Edge> candidates = byPayee.getOrDefault(current, Collections.emptyList());
            if (candidates.isEmpty()) break;

            Map<String, List<Edge>> groups = new LinkedHashMap<>();
            for (Edge e : candidates) {
                if (usedRows.contains(e.rowNo)) continue;
                groups.computeIfAbsent(e.payer, k -> new ArrayList<>()).add(e);
            }
            if (groups.isEmpty()) break;

            String chosenPayer = null;
            int minRow = Integer.MAX_VALUE;
            for (Map.Entry<String, List<Edge>> entry : groups.entrySet()) {
                int row = entry.getValue().get(0).rowNo;
                if (row < minRow) {
                    minRow = row;
                    chosenPayer = entry.getKey();
                }
            }

            for (String payer : groups.keySet()) {
                if (!payer.equals(chosenPayer)) {
                    queue.addLast(current);
                }
            }

            List<Edge> chosenEdges = groups.get(chosenPayer);
            for (int i = chosenEdges.size() - 1; i >= 0; i--) {
                Edge e = chosenEdges.get(i);
                usedRows.add(e.rowNo);
                chain.edges.add(e);
            }

            current = chosenPayer;
        }

        return chain;
    }

    private ChainWriteResult writeLinearChain(Sheet out,
                                              Chain chain,
                                              int rowStart,
                                              CellStyle nameStyle,
                                              CellStyle amtStyle) {
        if (chain.edges.isEmpty()) return new ChainWriteResult(0, 0);

        List<Edge> edges = new ArrayList<>(chain.edges);
        Collections.reverse(edges); // 头 -> 尾

        // 按供方→需方分组，合并同一对公司的多笔交易
        List<String> companies = new ArrayList<>();
        List<List<BigDecimal>> amountGroups = new ArrayList<>();
        
        companies.add(edges.get(0).payer); // 链头（最终付款方）
        
        String lastPayee = null;
        List<BigDecimal> currentAmounts = null;
        
        for (Edge e : edges) {
            if (!e.payee.equals(lastPayee)) {
                // 新的供方，创建新列
                if (currentAmounts != null) {
                    amountGroups.add(currentAmounts);
                }
                companies.add(e.payee);
                currentAmounts = new ArrayList<>();
                lastPayee = e.payee;
            }
            currentAmounts.add(e.amount);
        }
        if (currentAmounts != null) {
            amountGroups.add(currentAmounts);
        }

        // 写入公司名称（第一行）
        int nameRow = rowStart;
        for (int col = 0; col < companies.size(); col++) {
            writeNameCell(out, col, nameRow, companies.get(col), nameStyle);
        }

        // 写入金额（从第二行开始，可能多行）
        int maxAmounts = 0;
        for (int col = 0; col < amountGroups.size(); col++) {
            List<BigDecimal> amounts = amountGroups.get(col);
            maxAmounts = Math.max(maxAmounts, amounts.size());
            for (int i = 0; i < amounts.size(); i++) {
                writeAmountCell(out, col, rowStart + 1 + i, amounts.get(i), amtStyle);
            }
        }

        int totalRows = 1 + maxAmounts; // 名称行 + 金额行数
        return new ChainWriteResult(companies.size(), totalRows);
    }

    private void writeNameCell(Sheet out, int col, int row, String name, CellStyle style) {
        Row r = out.getRow(row);
        if (r == null) r = out.createRow(row);
        Cell c = r.getCell(col);
        if (c == null) c = r.createCell(col);
        c.setCellStyle(style);
        c.setCellValue(name == null ? "" : name);
    }

    private void writeAmountCell(Sheet out, int col, int row, BigDecimal amount, CellStyle style) {
        if (amount == null) return;
        Row r = out.getRow(row);
        if (r == null) r = out.createRow(row);
        Cell c = r.getCell(col);
        if (c == null) c = r.createCell(col);
        c.setCellStyle(style);
        c.setCellValue(amount.doubleValue());
    }

    private CellStyle createNameStyle(Workbook wb) {
        CellStyle s = wb.createCellStyle();
        s.setAlignment(HorizontalAlignment.LEFT);
        s.setVerticalAlignment(VerticalAlignment.CENTER);
        return s;
    }

    private CellStyle createAmountStyle(Workbook wb) {
        CellStyle s = wb.createCellStyle();
        s.setDataFormat(wb.createDataFormat().getFormat("#,##0.00"));
        s.setAlignment(HorizontalAlignment.LEFT);
        s.setVerticalAlignment(VerticalAlignment.CENTER);
        return s;
    }

    // =========================================================================
    // 数据读取
    // =========================================================================

    private List<Edge> readEdges(Sheet sheet) {
        List<Edge> edges = new ArrayList<>();
        if (sheet.getPhysicalNumberOfRows() < 2) return edges;

        Row headerRow = sheet.getRow(0);
        if (headerRow == null) return edges;

        List<String> headers = new ArrayList<>();
        for (Cell cell : headerRow) {
            headers.add(getCellStringValue(cell, null));
        }

        int payerIdx  = findHeaderIndex(headers, Arrays.asList("需方", "买方", "需方（提货单位）", "需方／提货单位"));
        int payeeIdx  = findHeaderIndex(headers, Arrays.asList("供方", "卖方", "供方/发货单位", "供方／发货单位"));
        int amountIdx = findHeaderIndex(headers, Collections.singletonList(config.getAmountColumn()));
        if (amountIdx < 0) {
            amountIdx = findHeaderIndex(headers, Arrays.asList("价税合计金额", "不含税金额", "结算金额"));
        }

        if (payerIdx < 0 || payeeIdx < 0 || amountIdx < 0) {
            log.warn("找不到必要列（需方/供方/金额），headers={}", headers);
            return edges;
        }

        for (int r = 1; r <= sheet.getLastRowNum(); r++) {
            Row row = sheet.getRow(r);
            if (row == null) continue;

            String payer  = getCellStringValue(row.getCell(payerIdx),  "需方").trim();
            String payee  = getCellStringValue(row.getCell(payeeIdx),  "供方").trim();
            String amtStr = getCellStringValue(row.getCell(amountIdx), config.getAmountColumn()).trim();

            if (payer.isEmpty() || payee.isEmpty()) continue;
            BigDecimal amount = parseAmount(amtStr);
            if (amount == null || amount.compareTo(BigDecimal.ZERO) <= 0) continue;

            edges.add(new Edge(payer, payee, amount, r + 1));
        }
        return edges;
    }

    // =========================================================================
    // 辅助方法
    // =========================================================================

    private static int findHeaderIndex(List<String> headers, List<String> candidates) {
        for (int i = 0; i < headers.size(); i++) {
            String h = headers.get(i) == null ? "" : headers.get(i).trim();
            for (String c : candidates) {
                if (c != null && !c.trim().isEmpty() && h.equals(c.trim())) return i;
            }
        }
        return -1;
    }

    private static BigDecimal parseAmount(String s) {
        if (s == null || s.trim().isEmpty()) return null;
        try {
            return new BigDecimal(s.trim().replace(",", "")).setScale(2, RoundingMode.HALF_UP);
        } catch (Exception ignored) {
            return null;
        }
    }

    private static Workbook openWorkbook(InputStream is, Path file) throws IOException {
        String name = file.getFileName().toString().toLowerCase();
        return name.endsWith(".xlsx") ? new XSSFWorkbook(is) : new HSSFWorkbook(is);
    }

    private static String sanitizeFileName(String name) {
        if (name == null) return "";
        return name.replaceAll("[\\\\/:*?\"<>|]", "_").trim();
    }

    private static String sanitizeSheetName(String name) {
        String safe = sanitizeFileName(name);
        return safe.length() > 31 ? safe.substring(0, 31) : safe;
    }

    private static final DateTimeFormatter DATE_FMT = DateTimeFormatter.ofPattern("yyyy/M/d");

    private String getCellStringValue(Cell cell, String columnHeader) {
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    try { return cell.getLocalDateTimeCellValue().toLocalDate().format(DATE_FMT); }
                    catch (Exception ignored) { return ""; }
                }
                return formatNumeric(cell.getNumericCellValue(), isAmountColumn(columnHeader) ? 2 : -1);
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                try {
                    if (DateUtil.isCellDateFormatted(cell)) {
                        return cell.getLocalDateTimeCellValue().toLocalDate().format(DATE_FMT);
                    }
                    return formatNumeric(cell.getNumericCellValue(), isAmountColumn(columnHeader) ? 2 : -1);
                } catch (Exception e) {
                    try { return cell.getStringCellValue(); }
                    catch (Exception ignored) { return cell.getCellFormula(); }
                }
            default:
                return "";
        }
    }

    private boolean isAmountColumn(String header) {
        if (header == null) return false;
        String h = header.trim();
        return h.equals("价税合计金额") || h.equals("不含税金额") || h.equals("增值税额")
                || h.equals(config.getAmountColumn());
    }

    private static String formatNumeric(double num, int decimals) {
        if (Double.isNaN(num) || Double.isInfinite(num)) return String.valueOf(num);
        if (decimals == 0 || (decimals < 0 && num == Math.floor(num) && Math.abs(num) < 1e15)) {
            return String.valueOf((long) num);
        }
        int scale = decimals >= 0 ? decimals : 6;
        return BigDecimal.valueOf(num).setScale(scale, RoundingMode.HALF_UP)
                         .stripTrailingZeros().toPlainString();
    }
}
