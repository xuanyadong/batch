package com.tjexcel.batch.service;

import com.tjexcel.batch.config.CustomerBelongingExtractConfig;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;

import java.io.BufferedInputStream;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.SimpleFileVisitor;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.nio.file.attribute.BasicFileAttributes;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Comparator;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;

@Service
public class CustomerBelongingExtractService {

    private static final Logger log = LoggerFactory.getLogger(CustomerBelongingExtractService.class);
    private static final List<Charset> READ_CHARSETS = Arrays.asList(
            StandardCharsets.UTF_8,
            Charset.forName("GBK"),
            Charset.forName("GB18030")
    );

    private final CustomerBelongingExtractConfig config;

    public CustomerBelongingExtractService(CustomerBelongingExtractConfig config) {
        this.config = config;
    }

    /**
     * 从“客属表”读取公司名，扫描源压缩包文件内容，将命中文件按sheet归类并输出新压缩包。
     *
     * @return 命中的文件数量
     */
    public int extract() throws IOException {
        Path dataPath = Paths.get(config.getDataPath()).toAbsolutePath();
        Path sourceZipPath = Paths.get(config.getSourceZipPath()).toAbsolutePath();
        Path outputDir = Paths.get(config.getOutputDir()).toAbsolutePath();

        if (!Files.exists(dataPath)) {
            throw new FileNotFoundException("客属数据表不存在: " + dataPath);
        }
        if (!Files.exists(sourceZipPath)) {
            throw new FileNotFoundException("源压缩包不存在: " + sourceZipPath);
        }

        Files.createDirectories(outputDir);
        Path workRoot = outputDir.resolve("customer_extract_work");
        recreateDirectory(workRoot);

        // 复制源压缩包，避免任何直接操作原文件。
        Path copiedZip = workRoot.resolve(sourceZipPath.getFileName().toString());
        Files.copy(sourceZipPath, copiedZip, StandardCopyOption.REPLACE_EXISTING);

        Map<String, Set<String>> companiesBySheet = readCompaniesBySheet(dataPath, config.getCompanyColumnName());
        if (companiesBySheet.isEmpty()) {
            log.warn("未从客属表读取到公司名，终止处理。数据表: {}", dataPath);
            return 0;
        }

        Path unzipDir = workRoot.resolve("unzipped");
        Path groupedDir = workRoot.resolve("grouped");
        Files.createDirectories(unzipDir);
        Files.createDirectories(groupedDir);
        ensureSheetFolders(groupedDir, companiesBySheet.keySet());

        int matchedCount = extractAndGroup(copiedZip, unzipDir, groupedDir, companiesBySheet);
        int zipCount = zipEachSheetFolder(groupedDir, outputDir);
        log.info("客属表文件提取完成，命中文件 {} 个，输出压缩包 {} 个，目录: {}",
                matchedCount, zipCount, outputDir);
        return matchedCount;
    }

    private Map<String, Set<String>> readCompaniesBySheet(Path dataPath, String companyColumnName) throws IOException {
        Map<String, Set<String>> result = new LinkedHashMap<>();
        try (InputStream is = Files.newInputStream(dataPath);
             Workbook workbook = openWorkbook(is, dataPath)) {
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                Sheet sheet = workbook.getSheetAt(i);
                if (sheet == null) {
                    continue;
                }
                String sheetName = sanitizeFolderName(sheet.getSheetName());
                Set<String> companies = readCompaniesFromSheet(sheet, companyColumnName);
                if (!companies.isEmpty()) {
                    result.put(sheetName, companies);
                    log.info("Sheet [{}] 读取到 {} 家公司", sheet.getSheetName(), companies.size());
                } else {
                    log.info("Sheet [{}] 未读取到公司名，跳过", sheet.getSheetName());
                }
            }
        }
        return result;
    }

    private Set<String> readCompaniesFromSheet(Sheet sheet, String companyColumnName) {
        Set<String> companies = new LinkedHashSet<>();
        if (sheet.getPhysicalNumberOfRows() < 2) {
            return companies;
        }

        int companyCol = resolveCompanyColumn(sheet.getRow(0), companyColumnName);
        if (companyCol < 0) {
            return companies;
        }

        for (int r = 1; r <= sheet.getLastRowNum(); r++) {
            Row row = sheet.getRow(r);
            if (row == null) {
                continue;
            }
            String company = getCellStringValue(row.getCell(companyCol)).trim();
            if (!company.isEmpty()) {
                companies.add(company);
            }
        }
        return companies;
    }

    private int resolveCompanyColumn(Row header, String companyColumnName) {
        if (header == null) {
            return -1;
        }
        String expected = defaultIfBlank(companyColumnName, "企业抬头").trim();
        short last = header.getLastCellNum();
        if (last < 0) {
            return -1;
        }
        for (int c = 0; c < last; c++) {
            String title = getCellStringValue(header.getCell(c)).trim();
            if (expected.equals(title)) {
                return c;
            }
        }
        return -1;
    }

    private int extractAndGroup(Path zipPath,
                                Path unzipDir,
                                Path groupedDir,
                                Map<String, Set<String>> companiesBySheet) throws IOException {
        int matched = 0;
        try (InputStream fis = Files.newInputStream(zipPath);
             BufferedInputStream bis = new BufferedInputStream(fis);
             ZipInputStream zis = new ZipInputStream(bis)) {
            ZipEntry entry;
            while ((entry = zis.getNextEntry()) != null) {
                if (entry.isDirectory()) {
                    zis.closeEntry();
                    continue;
                }
                String entryName = normalizeZipEntryName(entry.getName());
                byte[] content = readAllBytes(zis);

                Path extractedFile = unzipDir.resolve(entryName);
                Path parent = extractedFile.getParent();
                if (parent != null) {
                    Files.createDirectories(parent);
                }
                Files.write(extractedFile, content);

                Set<String> hitSheets = findHitSheets(content, entryName, companiesBySheet);
                if (!hitSheets.isEmpty()) {
                    for (String sheet : hitSheets) {
                        Path target = groupedDir.resolve(sheet).resolve(entryName);
                        Path targetParent = target.getParent();
                        if (targetParent != null) {
                            Files.createDirectories(targetParent);
                        }
                        Files.copy(extractedFile, target, StandardCopyOption.REPLACE_EXISTING);
                    }
                    matched++;
                    log.info("命中 {} 个sheet: {}", hitSheets.size(), entryName);
                }
                zis.closeEntry();
            }
        }
        return matched;
    }

    private Set<String> findHitSheets(byte[] content, Map<String, Set<String>> companiesBySheet) {
        return findHitSheets(content, "", companiesBySheet);
    }

    private Set<String> findHitSheets(byte[] content, String entryName, Map<String, Set<String>> companiesBySheet) {
        String text = extractSearchableText(content, entryName);
        Set<String> hitSheets = new LinkedHashSet<>();
        if (text.isEmpty()) {
            return hitSheets;
        }
        for (Map.Entry<String, Set<String>> item : companiesBySheet.entrySet()) {
            for (String company : item.getValue()) {
                if (company == null || company.trim().isEmpty()) {
                    continue;
                }
                if (text.contains(company.trim())) {
                    hitSheets.add(item.getKey());
                    break;
                }
            }
        }
        return hitSheets;
    }

    private String extractSearchableText(byte[] content, String entryName) {
        String lower = defaultIfBlank(entryName, "").toLowerCase(Locale.ROOT);
        try {
            if (lower.endsWith(".xls") || lower.endsWith(".xlsx")) {
                String excelText = extractExcelText(content);
                if (!excelText.isEmpty()) {
                    return excelText;
                }
            }
        } catch (Exception e) {
            log.warn("读取Excel文本失败，回退字节解码。entry={}，原因={}", entryName, e.getMessage());
        }
        return decodeToText(content);
    }

    private String extractExcelText(byte[] content) throws IOException {
        if (content == null || content.length == 0) {
            return "";
        }
        try (ByteArrayInputStream bis = new ByteArrayInputStream(content);
             Workbook workbook = WorkbookFactory.create(bis)) {
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                Sheet sheet = workbook.getSheetAt(i);
                if (sheet == null) {
                    continue;
                }
                for (int r = 0; r <= sheet.getLastRowNum(); r++) {
                    Row row = sheet.getRow(r);
                    if (row == null) {
                        continue;
                    }
                    short lastCell = row.getLastCellNum();
                    if (lastCell < 0) {
                        continue;
                    }
                    for (int c = 0; c < lastCell; c++) {
                        String cellText = getCellStringValue(row.getCell(c));
                        if (!cellText.isEmpty()) {
                            sb.append(cellText).append('\n');
                        }
                    }
                }
            }
            return sb.toString();
        }
    }

    private void ensureSheetFolders(Path groupedDir, Set<String> sheetNames) throws IOException {
        for (String sheet : sheetNames) {
            if (sheet == null || sheet.trim().isEmpty()) {
                continue;
            }
            Files.createDirectories(groupedDir.resolve(sheet));
        }
    }

    private String decodeToText(byte[] content) {
        if (content == null || content.length == 0) {
            return "";
        }

        String utf8 = new String(content, StandardCharsets.UTF_8);
        if (looksLikeReadableText(utf8)) {
            return utf8;
        }

        for (Charset charset : READ_CHARSETS) {
            String candidate = new String(content, charset);
            if (looksLikeReadableText(candidate)) {
                return candidate;
            }
        }
        return "";
    }

    private boolean looksLikeReadableText(String text) {
        if (text == null || text.trim().isEmpty()) {
            return false;
        }
        int checkLen = Math.min(text.length(), 4000);
        int printable = 0;
        for (int i = 0; i < checkLen; i++) {
            char ch = text.charAt(i);
            if (Character.isWhitespace(ch)
                    || (ch >= 32 && ch <= 126)
                    || (ch >= '\u4e00' && ch <= '\u9fa5')) {
                printable++;
            }
        }
        return printable >= checkLen * 0.55;
    }

    private void zipDirectory(Path sourceDir, Path targetZip) throws IOException {
        if (!Files.exists(sourceDir)) {
            throw new FileNotFoundException("待压缩目录不存在: " + sourceDir);
        }
        Path parent = targetZip.getParent();
        if (parent != null) {
            Files.createDirectories(parent);
        }
        try (ZipOutputStream zos = new ZipOutputStream(Files.newOutputStream(targetZip))) {
            List<Path> files = new ArrayList<>();
            Files.walk(sourceDir)
                    .filter(Files::isRegularFile)
                    .forEach(files::add);
            files.sort(Comparator.comparing(Path::toString));

            for (Path file : files) {
                String entryName = sourceDir.relativize(file).toString().replace('\\', '/');
                ZipEntry entry = new ZipEntry(entryName);
                zos.putNextEntry(entry);
                Files.copy(file, zos);
                zos.closeEntry();
            }
        }
    }

    private int zipEachSheetFolder(Path groupedDir, Path outputDir) throws IOException {
        if (!Files.exists(groupedDir)) {
            return 0;
        }
        List<Path> sheetDirs = new ArrayList<>();
        try (java.util.stream.Stream<Path> stream = Files.list(groupedDir)) {
            stream.filter(Files::isDirectory)
                    .sorted(Comparator.comparing(Path::toString))
                    .forEach(sheetDirs::add);
        }

        int count = 0;
        for (Path sheetDir : sheetDirs) {
            String suffix = resolveZipSuffixByFirstFile(sheetDir);
            String zipName = safeFileName(sheetDir.getFileName().toString() + suffix);
            if (!zipName.toLowerCase(Locale.ROOT).endsWith(".zip")) {
                zipName = zipName + ".zip";
            }
            Path outputZip = outputDir.resolve(zipName);
            zipDirectory(sheetDir, outputZip);
            count++;
            log.info("已输出sheet压缩包: {}", outputZip);
        }
        return count;
    }

    private String resolveZipSuffixByFirstFile(Path sheetDir) throws IOException {
        List<Path> files = new ArrayList<>();
        Files.walk(sheetDir)
                .filter(Files::isRegularFile)
                .sorted(Comparator.comparing(Path::toString))
                .forEach(files::add);
        if (files.isEmpty()) {
            return "";
        }
        String firstName = defaultIfBlank(files.get(0).getFileName().toString(), "");
        String baseName = firstName;
        int extIdx = firstName.lastIndexOf('.');
        if (extIdx > 0) {
            baseName = firstName.substring(0, extIdx);
        }
        if (baseName.startsWith("合同")) {
            return "合同";
        }
        if (baseName.startsWith("提单")) {
            return "提单";
        }
        return "";
    }

    private Workbook openWorkbook(InputStream is, Path file) throws IOException {
        String name = file.getFileName().toString().toLowerCase(Locale.ROOT);
        return name.endsWith(".xlsx") ? new XSSFWorkbook(is) : new HSSFWorkbook(is);
    }

    private String getCellStringValue(Cell cell) {
        if (cell == null) {
            return "";
        }
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                double n = cell.getNumericCellValue();
                if (n == Math.floor(n) && Math.abs(n) < 1e15) {
                    return String.valueOf((long) n);
                }
                return String.valueOf(n);
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                try {
                    CellType resultType = cell.getCachedFormulaResultType();
                    if (resultType == CellType.STRING) {
                        return cell.getStringCellValue();
                    }
                    if (resultType == CellType.NUMERIC) {
                        double n2 = cell.getNumericCellValue();
                        if (n2 == Math.floor(n2) && Math.abs(n2) < 1e15) {
                            return String.valueOf((long) n2);
                        }
                        return String.valueOf(n2);
                    }
                    if (resultType == CellType.BOOLEAN) {
                        return String.valueOf(cell.getBooleanCellValue());
                    }
                    return "";
                } catch (Exception ignored) {
                    return "";
                }
            default:
                return "";
        }
    }

    private String sanitizeFolderName(String name) {
        String safe = defaultIfBlank(name, "sheet")
                .replaceAll("[\\\\/:*?\"<>|]", "_")
                .trim();
        return safe.isEmpty() ? "sheet" : safe;
    }

    private String safeFileName(String fileName) {
        return defaultIfBlank(fileName, "result.zip")
                .replaceAll("[\\\\/:*?\"<>|]", "_")
                .trim();
    }

    private String defaultIfBlank(String value, String defaultValue) {
        if (value == null || value.trim().isEmpty()) {
            return defaultValue;
        }
        return value.trim();
    }

    private String normalizeZipEntryName(String name) {
        String normalized = defaultIfBlank(name, "unknown.bin").replace("\\", "/");
        while (normalized.startsWith("/")) {
            normalized = normalized.substring(1);
        }
        return normalized.replace("..", "_");
    }

    private byte[] readAllBytes(InputStream inputStream) throws IOException {
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        byte[] buffer = new byte[8192];
        int read;
        while ((read = inputStream.read(buffer)) != -1) {
            bos.write(buffer, 0, read);
        }
        return bos.toByteArray();
    }

    private void recreateDirectory(Path dir) throws IOException {
        if (Files.exists(dir)) {
            deleteDirectoryWithRetry(dir, 3, 200L);
        }
        Files.createDirectories(dir);
    }

    private void deleteDirectoryWithRetry(Path dir, int maxAttempts, long sleepMillis) throws IOException {
        IOException lastException = null;
        for (int i = 1; i <= maxAttempts; i++) {
            try {
                deleteDirectoryRecursively(dir);
                return;
            } catch (IOException e) {
                lastException = e;
                if (i < maxAttempts) {
                    log.warn("删除工作目录失败，准备重试。attempt={}/{}, dir={}, reason={}",
                            i, maxAttempts, dir, e.getMessage());
                    sleepQuietly(sleepMillis);
                }
            }
        }
        throw lastException == null ? new IOException("删除目录失败: " + dir) : lastException;
    }

    private void deleteDirectoryRecursively(Path dir) throws IOException {
        Files.walkFileTree(dir, new SimpleFileVisitor<Path>() {
            @Override
            public java.nio.file.FileVisitResult visitFile(Path file, BasicFileAttributes attrs) throws IOException {
                Files.deleteIfExists(file);
                return java.nio.file.FileVisitResult.CONTINUE;
            }

            @Override
            public java.nio.file.FileVisitResult postVisitDirectory(Path directory, IOException exc) throws IOException {
                if (exc != null) {
                    throw exc;
                }
                Files.deleteIfExists(directory);
                return java.nio.file.FileVisitResult.CONTINUE;
            }
        });
    }

    private void sleepQuietly(long millis) {
        try {
            Thread.sleep(millis);
        } catch (InterruptedException e) {
            Thread.currentThread().interrupt();
        }
    }
}
