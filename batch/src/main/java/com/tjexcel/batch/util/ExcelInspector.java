package com.tjexcel.batch.util;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.PrintStream;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

/**
 * 临时工具：读取 Excel 文件结构，用于分析数据填充方式
 * 运行: java -jar batch.jar --inspect 或 指定路径 --inspect --contract.data-path=... --contract.template-path=...
 */
public class ExcelInspector {

    private final PrintStream out;

    public ExcelInspector(PrintStream out) {
        this.out = out;
    }

    public static void run(Path dataPath, Path templatePath, Path filledPath, PrintStream out) throws Exception {
        ExcelInspector inspector = new ExcelInspector(out);
        out.println("========== 数据表 0126 第1-2行 A-H列 ==========");
        inspector.inspectSheet(dataPath, 0, 0, 1, 0, 8);
        out.println("\n========== 塑料合同模板 所有非空单元格 ==========");
        inspector.inspectAllNonEmpty(templatePath);
        out.println("\n========== 已填充合同 所有非空单元格 ==========");
        inspector.inspectAllNonEmpty(filledPath);
    }

    public static void main(String[] args) throws Exception {
        Path base = Paths.get("src/main/resources");
        Path data = base.resolve("0126.xlsx");
        Path template = base.resolve("塑料合同模板(2).xlsx");
        Path filled = base.resolve("合同-伊科东城-奥卓-XMYKAZ2026113.xls");
        run(data, template, filled, System.out);
    }

    void inspectSheet(Path file, int sheetIndex, int fromRow, int toRow, int fromCol, int toCol) throws Exception {
        try (InputStream is = Files.newInputStream(file);
             Workbook wb = file.toString().toLowerCase().endsWith("xlsx") ? new XSSFWorkbook(is) : new HSSFWorkbook(is)) {
            Sheet s = wb.getSheetAt(sheetIndex);
            for (int r = fromRow; r <= toRow; r++) {
                Row row = s.getRow(r);
                out.print("行" + (r + 1) + ": ");
                for (int c = fromCol; c < toCol; c++) {
                    String v = "";
                    if (row != null) {
                        Cell cell = row.getCell(c);
                        v = getVal(cell);
                    }
                    out.print("[" + v + "] ");
                }
                out.println();
            }
        }
    }

    void inspectAllNonEmpty(Path file) throws Exception {
        try (InputStream is = Files.newInputStream(file);
             Workbook wb = file.toString().toLowerCase().endsWith("xlsx") ? new XSSFWorkbook(is) : new HSSFWorkbook(is)) {
            for (int si = 0; si < wb.getNumberOfSheets(); si++) {
                Sheet s = wb.getSheetAt(si);
                out.println("--- Sheet: " + s.getSheetName() + " ---");
                int count = 0;
                for (Row row : s) {
                    for (Cell cell : row) {
                        String v = getVal(cell);
                        if (v != null && !v.trim().isEmpty()) {
                            String coord = toCoord(cell.getRowIndex(), cell.getColumnIndex());
                            out.println("  " + coord + ": " + (v.length() > 80 ? v.substring(0, 80) + "..." : v));
                            if (++count >= 80) break;
                        }
                    }
                    if (count >= 80) break;
                }
            }
        }
    }

    String getVal(Cell c) {
        if (c == null) return "";
        switch (c.getCellType()) {
            case STRING: return c.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(c)) return c.getLocalDateTimeCellValue().toString();
                double n = c.getNumericCellValue();
                return n == (long) n ? String.valueOf((long) n) : String.valueOf(n);
            case BOOLEAN: return String.valueOf(c.getBooleanCellValue());
            case FORMULA: return "FORMULA:" + c.getCellFormula();
            default: return "";
        }
    }

    String toCoord(int r, int c) {
        return (char) ('A' + c) + String.valueOf(r + 1);
    }
}
