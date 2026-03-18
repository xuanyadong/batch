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

                List<String> roots = resolveRoots(edges);
                if (roots.isEmpty()) {
                    log.warn("Sheet [{}] 未找到起点，跳过", sheet.getSheetName());
                    continue;
                }

                String outName = "资金流-" + sanitizeFileName(sheet.getSheetName()) + ".xlsx";
                Path outFile = outputDir.resolve(outName);

                try (XSSFWorkbook outWb = new XSSFWorkbook()) {
                    String sheetName = sanitizeSheetName(sheet.getSheetName());
                    if (sheetName.isEmpty()) sheetName = "资金流";
                    Sheet outSheet = outWb.createSheet(sheetName);
                    writeSheet(outWb, outSheet, edges, roots);
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

    /** 一条有向交易记录（有向边）。rowNo 为数据表原始行号（1-based），用于排序和唯一性。 */
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

    // =========================================================================
    // Excel 写入
    // =========================================================================

    private void writeSheet(Workbook wb, Sheet out, List<Edge> allEdges, List<String> roots) {
        CellStyle nameStyle = createNameStyle(wb);
        CellStyle amtStyle  = createAmountStyle(wb);

        // 建图：需方 -> 出边列表，按 rowNo 升序（规则 5.1）
        Map<String, List<Edge>> graph = new LinkedHashMap<>();
        for (Edge e : allEdges) {
            graph.computeIfAbsent(e.payer, k -> new ArrayList<>()).add(e);
        }
        graph.values().forEach(list -> list.sort(Comparator.comparingInt(e -> e.rowNo)));
        // 入边：供方 -> 入边列表，按 rowNo 升序（规则 5.2）
        Map<String, List<Edge>> incoming = new LinkedHashMap<>();
        for (Edge e : allEdges) {
            incoming.computeIfAbsent(e.payee, k -> new ArrayList<>()).add(e);
        }
        incoming.values().forEach(list -> list.sort(Comparator.comparingInt(e -> e.rowNo)));

        Set<Integer> usedRows = new LinkedHashSet<>(); // 规则 7.1：每条交易只出现一次

        int rowCursor = 0;
        int maxCol    = 0;
        int[] maxColRef = {0};

        // 1) 先按配置起点生成链
        for (String root : roots) {
            rowCursor = writeChain(out, root, graph, new LinkedHashSet<>(),
                    0, rowCursor, nameStyle, amtStyle, maxColRef, usedRows, incoming);
            maxCol = Math.max(maxCol, maxColRef[0]);
            rowCursor++; // 规则 10：链间空行
        }

        // 2) 对未使用的交易补链：按原始行号升序，每条以付款方为起点独立成链
        List<Edge> remaining = new ArrayList<>();
        for (Edge e : allEdges) {
            if (!usedRows.contains(e.rowNo)) remaining.add(e);
        }
        remaining.sort(Comparator.comparingInt(e -> e.rowNo));
        for (Edge e : remaining) {
            rowCursor = writeChain(out, e.payer, graph, new LinkedHashSet<>(),
                    0, rowCursor, nameStyle, amtStyle, maxColRef, usedRows, incoming);
            maxCol = Math.max(maxCol, maxColRef[0]);
            rowCursor++; // 链间空行
        }

        for (int c = 0; c <= maxCol; c++) {
            out.autoSizeColumn(c);
            int w = out.getColumnWidth(c);
            int adjusted = (int) Math.min((long) w * 3L / 2L, 255L * 256L);
            out.setColumnWidth(c, Math.max(20 * 256, adjusted));
        }
    }

    /**
     * 递归写链：返回写完后的下一可用行。
     *
     * 规则要点：
     * - 每笔交易占两行：名称行 + 金额行（金额写在付款方列）
     * - 分叉在付款方列下方展开（规则 9）
     * - 叶节点只写名称行
     */
    private int writeChain(Sheet out,
                           String company,
                           Map<String, List<Edge>> graph,
                           LinkedHashSet<String> path,
                           int col,
                           int rowStart,
                           CellStyle nameStyle,
                           CellStyle amtStyle,
                           int[] maxCol,
                           Set<Integer> usedRows,
                           Map<String, List<Edge>> incoming) {
        maxCol[0] = Math.max(maxCol[0], col);

        List<Edge> edges = graph.getOrDefault(company, Collections.emptyList());
        boolean hasUnused = false;
        for (Edge e : edges) {
            if (!usedRows.contains(e.rowNo)) {
                hasUnused = true;
                break;
            }
        }

        // 叶节点或闭环：只写公司名
        if (path.contains(company) || !hasUnused) {
            writeNameCell(out, col, rowStart, company, nameStyle);
            return rowStart + 1;
        }

        LinkedHashSet<String> nextPath = new LinkedHashSet<>(path);
        nextPath.add(company);

        // 链首优先写一行公司名，防止链首缺失
        writeNameCell(out, col, rowStart, company, nameStyle);

        int rowCursor = rowStart;
        for (Edge edge : edges) {
            if (usedRows.contains(edge.rowNo)) continue;
            if (!isIncomingOrderOk(edge, usedRows, incoming)) continue;
            usedRows.add(edge.rowNo);

            int nameRow = rowCursor;
            int amtRow  = rowCursor + 1;

            // 付款方公司名（每笔交易一行都写，避免出现“只有金额没有公司名”）
            writeNameCell(out, col, nameRow, company, nameStyle);
            // 金额（写在付款方列）
            writeAmountCell(out, col, amtRow, edge.amount, amtStyle);

            // 收款方在右侧继续展开
            int childEnd = writeChain(out, edge.payee, graph, nextPath,
                    col + 1, nameRow, nameStyle, amtStyle, maxCol, usedRows, incoming);

            rowCursor = Math.max(amtRow + 1, childEnd);
        }

        return rowCursor;
    }

    /**
     * 规则 5.2：同一供方的入边按行号升序优先展示。
     * 若当前入边前面还有未使用的更小行号入边，则先跳过当前入边。
     */
    private boolean isIncomingOrderOk(Edge edge,
                                      Set<Integer> usedRows,
                                      Map<String, List<Edge>> incoming) {
        List<Edge> inList = incoming.get(edge.payee);
        if (inList == null || inList.isEmpty()) return true;
        for (Edge e : inList) {
            if (e.rowNo >= edge.rowNo) break;
            if (!usedRows.contains(e.rowNo)) return false;
        }
        return true;
    }

    // =========================================================================
    // 单元格写入
    // =========================================================================

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

        int payerIdx  = findHeaderIndex(headers, Arrays.asList("需方", "买方", "需方/提货单位", "需方／提货单位"));
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

    /** 解析起点：优先用配置，配置为空则回退到入度为 0 的付款方。 */
    private List<String> resolveRoots(List<Edge> edges) {
        List<String> configured = config.getRootCompanies();
        if (configured != null) {
            List<String> valid = new ArrayList<>();
            for (String s : configured) {
                if (s != null && !s.trim().isEmpty()) valid.add(s.trim());
            }
            if (!valid.isEmpty()) return valid;
        }

        Set<String> payees = new HashSet<>();
        for (Edge e : edges) payees.add(e.payee);

        List<String> roots = new ArrayList<>();
        LinkedHashSet<String> seen = new LinkedHashSet<>();
        for (Edge e : edges) {
            if (seen.add(e.payer) && !payees.contains(e.payer)) roots.add(e.payer);
        }
        if (roots.isEmpty() && !seen.isEmpty()) roots.add(seen.iterator().next());
        return roots;
    }

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
