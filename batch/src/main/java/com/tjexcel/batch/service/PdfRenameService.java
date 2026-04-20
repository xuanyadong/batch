package com.tjexcel.batch.service;

import com.tjexcel.batch.config.PdfRenameConfig;
import nu.pattern.OpenCV;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.rendering.ImageType;
import org.apache.pdfbox.rendering.PDFRenderer;
import org.opencv.core.Core;
import org.opencv.core.CvType;
import org.opencv.core.Mat;
import org.opencv.core.Scalar;
import org.opencv.core.Size;
import org.opencv.imgproc.Imgproc;
import org.opencv.photo.Photo;
import net.sourceforge.pinyin4j.PinyinHelper;
import net.sourceforge.pinyin4j.format.HanyuPinyinOutputFormat;
import net.sourceforge.pinyin4j.format.HanyuPinyinToneType;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.geom.AffineTransform;
import java.awt.image.BufferedImage;
import java.awt.image.DataBufferByte;
import java.awt.image.WritableRaster;
import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.nio.charset.StandardCharsets;
import java.nio.file.*;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;
import java.util.stream.Stream;

@Service
public class PdfRenameService {

    private static final Logger log = LoggerFactory.getLogger(PdfRenameService.class);

    private static final Pattern CONTRACT_NO = Pattern.compile("合同编号\\s*[：:;；]\\s*([A-Za-z0-9\\-_/]+)");
    private static final Pattern BILL_NO = Pattern.compile("(提单编号|编号)\\s*[：:;；]\\s*([A-Za-z0-9\\-_/]+)");

    private final PdfRenameConfig config;
    private static volatile boolean OPEN_CV_READY = false;
    private static volatile boolean OPEN_CV_INIT_ATTEMPTED = false;

    public PdfRenameService(PdfRenameConfig config) {
        this.config = config;
    }

    public int renameByOcr() throws IOException {
        Path root = Paths.get(config.getScanDir()).toAbsolutePath();
        if (!Files.exists(root) || !Files.isDirectory(root)) {
            throw new IOException("PDF扫描目录不存在或不是目录: " + root);
        }

        Path tesseractExe = Paths.get(config.getTesseractExePath()).toAbsolutePath();
        if (!Files.exists(tesseractExe)) {
            throw new IOException("Tesseract 可执行文件不存在: " + tesseractExe);
        }

        List<Path> pdfFiles;
        try (Stream<Path> stream = config.isRecursive() ? Files.walk(root) : Files.list(root)) {
            pdfFiles = stream
                    .filter(Files::isRegularFile)
                    .filter(p -> p.getFileName().toString().toLowerCase(Locale.ROOT).endsWith(".pdf"))
                    .sorted(Comparator.comparing(Path::toString))
                    .collect(Collectors.toList());
        }

        if (pdfFiles.isEmpty()) {
            log.info("目录中未找到PDF文件: {}", root);
            return 0;
        }

        int success = 0;
        int failed = 0;
        for (Path pdf : pdfFiles) {
            try {
                String newName = buildNewFileName(pdf, tesseractExe);
                if (newName == null || newName.trim().isEmpty()) {
                    failed++;
                    log.warn("未识别出可用命名字段，跳过: {}", pdf.getFileName());
                    continue;
                }

                String safeName = sanitizeFileName(newName);
                if (!safeName.toLowerCase(Locale.ROOT).endsWith(".pdf")) {
                    safeName = safeName + ".pdf";
                }
                Path target = ensureUniqueTarget(pdf, safeName);

                if (target.equals(pdf)) {
                    log.info("文件名无需变更: {}", pdf.getFileName());
                    success++;
                    continue;
                }

                if (config.isDryRun()) {
                    log.info("[DRY-RUN] {} -> {}", pdf.getFileName(), target.getFileName());
                } else {
                    Files.move(pdf, target);
                    log.info("重命名成功: {} -> {}", pdf.getFileName(), target.getFileName());
                }
                success++;
            } catch (Exception ex) {
                failed++;
                log.error("处理失败: {}，原因: {}", pdf, ex.getMessage(), ex);
            }
        }

        log.info("PDF重命名完成。总数={}，成功={}，失败={}，dryRun={}",
                pdfFiles.size(), success, failed, config.isDryRun());
        return success;
    }

    private String buildNewFileName(Path pdf, Path tesseractExe) throws IOException {
        try (PDDocument document = PDDocument.load(pdf.toFile())) {
            int pageCount = document.getNumberOfPages();

            if (pageCount > 1) {
                String ocrText = runOcrForDocument(document, tesseractExe, pageCount);
                String normalized = normalizeText(ocrText);
                String partyA = extractValue(normalized, "甲方", "甲方（委托方）", "委托方");
                String partyB = extractValue(normalized, "乙方", "乙方（受托方）", "受托方");
                if (isBlank(partyA)) {
                    partyA = extractValueLoose(normalized, "甲方", "委托方");
                }
                if (isBlank(partyB)) {
                    partyB = extractValueLoose(normalized, "乙方", "受托方");
                }
                if (isBlank(partyA) || isBlank(partyB)) {
                    List<String> companies = extractCompanyCandidates(normalized);
                    if (isBlank(partyA) && companies.size() > 0) partyA = companies.get(0);
                    if (isBlank(partyB) && companies.size() > 1) partyB = companies.get(1);
                }
                if (isBlank(partyA) || isBlank(partyB)) {
                    log.warn("多页PDF未提取到甲乙方。文件={}", pdf.getFileName());
                    log.debug("OCR片段: {}", normalized.substring(0, Math.min(300, normalized.length())).replaceAll("\\s+", " "));
                    return null;
                }
                return partyA + "-" + partyB + "-仓储合同.pdf";
            }

            // Single-page docs: run a lightweight single OCR once on original image.
            String singlePageText = runSinglePageRawOcr(document, tesseractExe, 0, config.getPsm());
            String normalized = normalizeText(singlePageText);

            if (normalized.contains("购销合同")) {
                String supplier = extractValue(normalized, "供方");
                String buyer = extractValue(normalized, "需方");
                String contractNo = extractContractNo(normalized);
                supplier = cleanCompanyName(supplier);
                buyer = cleanCompanyName(buyer);
                if (isBlank(supplier) || isBlank(buyer) || isBlank(contractNo)) {
                    return null;
                }
                return supplier + "-" + buyer + "-合同-" + contractNo + ".pdf";
            }

            if (normalized.contains("过户单")) {
                TransferSlipFields roiFields = extractTransferSlipFields(document, tesseractExe);
                String shipper = roiFields.shipper;
                String consignee = roiFields.consignee;
                String billNo = roiFields.billNo;

                if (!isLikelyCompanyName(shipper)) {
                    shipper = extractValue(normalized, "发货单位", "发货单住", "发货单位", "发貨单位");
                }
                if (!isLikelyCompanyName(shipper)) {
                    shipper = extractValueLoose(normalized, "发货单位", "发货单住", "发货单位", "发貨单位");
                }
                if (!isLikelyCompanyName(shipper)) {
                    shipper = extractTransferPartyLineFallback(normalized, true);
                }
                if (!isLikelyCompanyName(shipper)) {
                    shipper = extractCompanyAfterLabel(normalized, "发货单位", "发貨单位", "发货单住");
                }
                if (!isLikelyCompanyName(consignee)) {
                    consignee = extractValue(normalized, "提货单位", "提货单住", "提货单位", "提貨单位");
                }
                if (!isLikelyCompanyName(consignee)) {
                    consignee = extractValueLoose(normalized, "提货单位", "提货单住", "提货单位", "提貨单位");
                }
                if (!isLikelyCompanyName(consignee)) {
                    consignee = extractTransferPartyLineFallback(normalized, false);
                }
                if (!isLikelyCompanyName(consignee)) {
                    consignee = extractCompanyAfterLabel(normalized, "提货单位", "提貨单位", "提货单住");
                }
                if (!isLikelyBillNo(billNo)) {
                    billNo = extractBillNo(normalized);
                }

                shipper = cleanCompanyName(shipper);
                consignee = cleanCompanyName(consignee);
                billNo = normalizeBillNo(billNo);

                // 避免从“请过户至...”备注区误识别公司：不再做全页公司候选硬推断。
                // 优先保证正确性：字段不足时宁可跳过，不拼凑错误名称。

                // 提单号一旦已是合法格式，禁止再按公司名修复前缀，避免误改编号。
                if (!isLikelyBillNo(billNo)) {
                    billNo = repairBillNoByCompanies(billNo, shipper, consignee);
                }

                if (isBlank(shipper) || isBlank(consignee) || isBlank(billNo)) {
                    log.warn("过户单字段提取不足。文件={} shipper='{}' consignee='{}' billNo='{}'",
                            pdf.getFileName(), shipper, consignee, billNo);
                    return null;
                }
                return shipper + "-" + consignee + "-提单-" + billNo + ".pdf";
            }

            return null;
        }
    }

    private String runSinglePageRawOcr(PDDocument document, Path tesseractExe, int pageIndex, int psm) throws IOException {
        if (document.getNumberOfPages() <= pageIndex) {
            return "";
        }
        PDFRenderer renderer = new PDFRenderer(document);
        BufferedImage pageImage = renderer.renderImageWithDPI(pageIndex, config.getDpi(), ImageType.RGB);
        return ocrImage(pageImage, tesseractExe, psm);
    }

    private String runOcrForDocument(PDDocument document, Path tesseractExe, int pageCount) throws IOException {
        PDFRenderer renderer = new PDFRenderer(document);
        int maxPages = config.getMaxPages() <= 0 ? pageCount : Math.min(config.getMaxPages(), pageCount);

        StringBuilder fullText = new StringBuilder();
        for (int i = 0; i < maxPages; i++) {
            BufferedImage pageImage = renderer.renderImageWithDPI(i, config.getDpi(), ImageType.RGB);
            BufferedImage preprocessed = preprocessForStampNoise(pageImage);
            String pageText;
            if (config.isFastMode()) {
                pageText = ocrImage(preprocessed, tesseractExe, config.getPsm());
            } else {
                BufferedImage bestDeskew = pickBestDeskewImage(preprocessed, tesseractExe);
                pageText = ocrImage(bestDeskew, tesseractExe, config.getPsm());
            }
            fullText.append('\n').append(pageText);
        }
        return fullText.toString();
    }

    /**
     * Field-level OCR for transfer-slip layout to avoid full-page noise.
     */
    private TransferSlipFields extractTransferSlipFields(PDDocument document, Path tesseractExe) throws IOException {
        if (document.getNumberOfPages() <= 0) {
            return TransferSlipFields.empty();
        }
        PDFRenderer renderer = new PDFRenderer(document);
        int dpi = Math.max(config.getDpi(), 260);
        BufferedImage page = renderer.renderImageWithDPI(0, dpi, ImageType.RGB);
        List<BufferedImage> variants = buildTransferSlipVariants(page);

        // Ratios tuned for current transfer-slip template (first page).
        String billRaw = ocrRegionBestOfVariants(variants, tesseractExe, 0.12, 0.145, 0.34, 0.075, 7, true);
        String shipperRaw = ocrRegionBestOfVariants(variants, tesseractExe, 0.14, 0.215, 0.34, 0.075, 7, false);
        String consigneeRaw = ocrRegionBestOfVariants(variants, tesseractExe, 0.63, 0.215, 0.33, 0.075, 7, false);

        String billNo = normalizeBillNo(extractBillNo(normalizeText(billRaw)));
        if (!isLikelyBillNo(billNo)) {
            billNo = normalizeBillNo(billRaw);
        }
        String shipper = cleanCompanyName(shipperRaw);
        String consignee = cleanCompanyName(consigneeRaw);

        return new TransferSlipFields(shipper, consignee, billNo);
    }

    private List<BufferedImage> buildTransferSlipVariants(BufferedImage page) {
        List<BufferedImage> variants = new ArrayList<>();
        variants.add(page); // 原图：有时盖章压字时原图反而更可读
        BufferedImage cvVariant = preprocessWithOpenCvForStampNoise(page);
        if (cvVariant != null) {
            variants.add(cvVariant);
        }
        variants.add(preprocessForStampNoise(page));
        variants.add(preprocessForStampNoiseAggressive(page));
        return variants;
    }

    private BufferedImage preprocessWithOpenCvForStampNoise(BufferedImage src) {
        if (!ensureOpenCvLoaded()) {
            return null;
        }
        Mat bgr = null;
        Mat hsv = null;
        Mat mask1 = null;
        Mat mask2 = null;
        Mat redMask = null;
        Mat kernel = null;
        Mat inpainted = null;
        Mat gray = null;
        Mat binary = null;
        try {
            bgr = bufferedImageToMatBgr(src);
            hsv = new Mat();
            Imgproc.cvtColor(bgr, hsv, Imgproc.COLOR_BGR2HSV);

            mask1 = new Mat();
            mask2 = new Mat();
            // 红色在HSV中跨0度，分两段取掩膜
            Core.inRange(hsv, new Scalar(0, 60, 40), new Scalar(15, 255, 255), mask1);
            Core.inRange(hsv, new Scalar(160, 60, 40), new Scalar(180, 255, 255), mask2);
            redMask = new Mat();
            Core.bitwise_or(mask1, mask2, redMask);

            // 适度膨胀，让盖章边缘也参与修复
            kernel = Imgproc.getStructuringElement(Imgproc.MORPH_ELLIPSE, new Size(3, 3));
            Imgproc.dilate(redMask, redMask, kernel);

            // 使用图像修复而非直接涂白，尽量保留压章下文字结构
            inpainted = new Mat();
            Photo.inpaint(bgr, redMask, inpainted, 3.0, Photo.INPAINT_TELEA);

            gray = new Mat();
            Imgproc.cvtColor(inpainted, gray, Imgproc.COLOR_BGR2GRAY);
            binary = new Mat();
            Imgproc.adaptiveThreshold(gray, binary, 255,
                    Imgproc.ADAPTIVE_THRESH_GAUSSIAN_C, Imgproc.THRESH_BINARY, 35, 8);
            return matToBufferedImage(binary);
        } catch (Exception e) {
            log.warn("OpenCV去章预处理失败，回退到原有流程: {}", e.getMessage());
            return null;
        } finally {
            safeRelease(binary);
            safeRelease(gray);
            safeRelease(inpainted);
            safeRelease(kernel);
            safeRelease(redMask);
            safeRelease(mask2);
            safeRelease(mask1);
            safeRelease(hsv);
            safeRelease(bgr);
        }
    }

    private boolean ensureOpenCvLoaded() {
        if (OPEN_CV_READY) {
            return true;
        }
        if (OPEN_CV_INIT_ATTEMPTED) {
            return false;
        }
        synchronized (PdfRenameService.class) {
            if (OPEN_CV_READY) {
                return true;
            }
            if (OPEN_CV_INIT_ATTEMPTED) {
                return false;
            }
            OPEN_CV_INIT_ATTEMPTED = true;
            try {
                OpenCV.loadLocally();
                OPEN_CV_READY = true;
                log.info("OpenCV native loaded for PDF stamp preprocessing");
            } catch (Throwable t) {
                OPEN_CV_READY = false;
                log.warn("OpenCV native load failed, fallback to Java preprocessing: {}", t.getMessage());
            }
        }
        return OPEN_CV_READY;
    }

    private Mat bufferedImageToMatBgr(BufferedImage image) throws IOException {
        BufferedImage rgb = image;
        if (image.getType() != BufferedImage.TYPE_3BYTE_BGR) {
            BufferedImage converted = new BufferedImage(image.getWidth(), image.getHeight(), BufferedImage.TYPE_3BYTE_BGR);
            Graphics2D g = converted.createGraphics();
            try {
                g.setColor(Color.WHITE);
                g.fillRect(0, 0, converted.getWidth(), converted.getHeight());
                g.drawImage(image, 0, 0, null);
            } finally {
                g.dispose();
            }
            rgb = converted;
        }
        byte[] data = ((DataBufferByte) rgb.getRaster().getDataBuffer()).getData();
        Mat mat = new Mat(rgb.getHeight(), rgb.getWidth(), CvType.CV_8UC3);
        mat.put(0, 0, data);
        return mat;
    }

    private BufferedImage matToBufferedImage(Mat mat) throws IOException {
        Mat output = mat;
        if (mat.channels() == 1) {
            output = new Mat();
            Imgproc.cvtColor(mat, output, Imgproc.COLOR_GRAY2BGR);
        }
        int width = output.width();
        int height = output.height();
        int channels = output.channels();
        byte[] source = new byte[width * height * channels];
        output.get(0, 0, source);

        BufferedImage image = new BufferedImage(width, height, BufferedImage.TYPE_3BYTE_BGR);
        byte[] target = ((DataBufferByte) image.getRaster().getDataBuffer()).getData();
        System.arraycopy(source, 0, target, 0, source.length);

        if (output != mat) {
            output.release();
        }
        return image;
    }

    private void safeRelease(Mat mat) {
        if (mat != null) {
            mat.release();
        }
    }

    private String ocrRegionBestOfVariants(List<BufferedImage> variants, Path tesseractExe,
                                           double rx, double ry, double rw, double rh, int psm,
                                           boolean preferBillNo) throws IOException {
        String bestText = "";
        int bestScore = Integer.MIN_VALUE;
        for (BufferedImage variant : variants) {
            String text = ocrRegion(variant, tesseractExe, rx, ry, rw, rh, psm);
            int score = preferBillNo ? scoreBillNoText(text) : scoreCompanyText(text);
            if (score > bestScore) {
                bestScore = score;
                bestText = text;
            }
        }
        return bestText;
    }

    private int scoreBillNoText(String text) {
        String v = normalizeBillNo(text);
        if (isBlank(v)) return 0;
        int score = 0;
        if (v.contains("-")) score += 5;
        if (v.matches(".*\\d{8}.*")) score += 10;
        if (v.matches(".*\\d{3}$")) score += 8;
        score += Math.min(v.length(), 30);
        if (isLikelyBillNo(v)) score += 20;
        return score;
    }

    private int scoreCompanyText(String text) {
        String c = cleanCompanyName(text);
        if (isBlank(c)) return 0;
        int score = Math.min(c.length(), 40);
        if (c.contains("公司")) score += 10;
        if (c.contains("有限责任公司") || c.contains("有限公司")) score += 8;
        int zh = countChineseChars(c);
        score += Math.min(zh, 15);
        return score;
    }

    private String ocrRegion(BufferedImage image, Path tesseractExe,
                             double rx, double ry, double rw, double rh, int psm) throws IOException {
        int w = image.getWidth();
        int h = image.getHeight();
        int x = clamp((int) Math.round(w * rx), 0, w - 1);
        int y = clamp((int) Math.round(h * ry), 0, h - 1);
        int cw = clamp((int) Math.round(w * rw), 1, w - x);
        int ch = clamp((int) Math.round(h * rh), 1, h - y);
        BufferedImage crop = image.getSubimage(x, y, cw, ch);
        BufferedImage enlarged = scale(crop, 2.0);
        String text = ocrImage(enlarged, tesseractExe, psm);
        if (isBlank(text)) {
            text = ocrImage(crop, tesseractExe, psm);
        }
        return normalizeText(text);
    }

    private BufferedImage scale(BufferedImage src, double factor) {
        if (factor <= 1) return src;
        int w = Math.max(1, (int) Math.round(src.getWidth() * factor));
        int h = Math.max(1, (int) Math.round(src.getHeight() * factor));
        BufferedImage dst = new BufferedImage(w, h, src.getType());
        Graphics2D g = dst.createGraphics();
        try {
            g.setRenderingHint(RenderingHints.KEY_INTERPOLATION, RenderingHints.VALUE_INTERPOLATION_BILINEAR);
            g.setColor(Color.WHITE);
            g.fillRect(0, 0, w, h);
            g.drawImage(src, 0, 0, w, h, null);
        } finally {
            g.dispose();
        }
        return dst;
    }

    /**
     * Red-seal-friendly preprocessing:
     * 1) suppress red-like pixels to white
     * 2) grayscale
     * 3) median denoise (3x3)
     * 4) Otsu threshold binarization
     */
    private BufferedImage preprocessForStampNoise(BufferedImage src) {
        int w = src.getWidth();
        int h = src.getHeight();

        BufferedImage gray = new BufferedImage(w, h, BufferedImage.TYPE_BYTE_GRAY);
        WritableRaster grayRaster = gray.getRaster();
        for (int y = 0; y < h; y++) {
            for (int x = 0; x < w; x++) {
                int rgb = src.getRGB(x, y);
                int r = (rgb >> 16) & 0xFF;
                int g = (rgb >> 8) & 0xFF;
                int b = rgb & 0xFF;
                int out;
                boolean likelyRedSeal = (r > 110) && (r > (int) (g * 1.2)) && (r > (int) (b * 1.2));
                if (likelyRedSeal) {
                    out = 255;
                } else {
                    out = (int) (0.299 * r + 0.587 * g + 0.114 * b);
                }
                grayRaster.setSample(x, y, 0, out);
            }
        }

        BufferedImage denoised = medianBlur3x3(gray);
        return otsuBinarize(denoised);
    }

    /**
     * 更激进的去章：优先在字段区提升红章去除力度，适用于提单编号/发货单位被公章覆盖。
     */
    private BufferedImage preprocessForStampNoiseAggressive(BufferedImage src) {
        int w = src.getWidth();
        int h = src.getHeight();

        BufferedImage gray = new BufferedImage(w, h, BufferedImage.TYPE_BYTE_GRAY);
        WritableRaster grayRaster = gray.getRaster();
        for (int y = 0; y < h; y++) {
            for (int x = 0; x < w; x++) {
                int rgb = src.getRGB(x, y);
                int r = (rgb >> 16) & 0xFF;
                int g = (rgb >> 8) & 0xFF;
                int b = rgb & 0xFF;
                int out;
                boolean likelyRedSeal = (r > 85) && (r > g + 20) && (r > b + 20);
                if (likelyRedSeal) {
                    out = 255;
                } else {
                    out = (int) (0.299 * r + 0.587 * g + 0.114 * b);
                }
                grayRaster.setSample(x, y, 0, out);
            }
        }

        BufferedImage denoised = medianBlur3x3(gray);
        return otsuBinarize(denoised);
    }

    private BufferedImage pickBestDeskewImage(BufferedImage src, Path tesseractExe) throws IOException {
        double[] angles = new double[]{-1.2, -0.6, 0, 0.6, 1.2};
        BufferedImage best = src;
        int bestScore = Integer.MIN_VALUE;
        for (double angle : angles) {
            BufferedImage rotated = rotate(src, angle);
            String text = ocrImage(rotated, tesseractExe, config.getPsm());
            int score = scoreText(text);
            if (score > bestScore) {
                bestScore = score;
                best = rotated;
            }
        }
        return best;
    }

    private int scoreText(String text) {
        if (text == null) return 0;
        int len = text.replaceAll("\\s+", "").length();
        int keywordHit = 0;
        for (String kw : new String[]{"购销合同", "过户单", "供方", "需方", "合同编号", "发货单位", "提货单位", "甲方", "乙方"}) {
            if (text.contains(kw)) keywordHit += 20;
        }
        return len + keywordHit;
    }

    private String ocrImage(BufferedImage image, Path tesseractExe, int psm) throws IOException {
        Path tempFile = Files.createTempFile("ocr-pdf-", ".png");
        try {
            ImageIO.write(image, "png", tempFile.toFile());
            ProcessBuilder pb = new ProcessBuilder(
                    tesseractExe.toString(),
                    tempFile.toString(),
                    "stdout",
                    "-l",
                    config.getLanguage(),
                    "--oem",
                    "1",
                    "--psm",
                    String.valueOf(psm)
            );
            pb.redirectErrorStream(true);
            Process process = pb.start();
            String output;
            try (BufferedReader reader = new BufferedReader(
                    new InputStreamReader(process.getInputStream(), StandardCharsets.UTF_8))) {
                StringBuilder sb = new StringBuilder();
                String line;
                while ((line = reader.readLine()) != null) {
                    sb.append(line).append('\n');
                }
                output = sb.toString();
            }
            try {
                int exitCode = process.waitFor();
                if (exitCode != 0) {
                    log.warn("Tesseract返回非0退出码: {}", exitCode);
                }
            } catch (InterruptedException e) {
                Thread.currentThread().interrupt();
                throw new IOException("OCR进程被中断", e);
            }
            return output;
        } finally {
            try {
                Files.deleteIfExists(tempFile);
            } catch (Exception ignored) {
            }
        }
    }

    private String normalizeText(String text) {
        if (text == null) return "";
        return text.replace('\u3000', ' ')
                .replace("\r\n", "\n")
                .replace('\r', '\n');
    }

    private String extractContractNo(String text) {
        Matcher matcher = CONTRACT_NO.matcher(text);
        if (matcher.find()) {
            return cleanupValue(matcher.group(1));
        }
        return "";
    }

    private String extractBillNo(String text) {
        Matcher matcher = BILL_NO.matcher(text);
        if (matcher.find()) {
            String value = matcher.groupCount() >= 2 ? matcher.group(2) : matcher.group(1);
            return cleanupValue(value);
        }
        return "";
    }

    private String extractValue(String text, String... labels) {
        if (isBlank(text) || labels == null || labels.length == 0) return "";
        List<String> boundaryLabels = new ArrayList<>();
        boundaryLabels.add("合同编号");
        boundaryLabels.add("签约时间");
        boundaryLabels.add("提单编号");
        boundaryLabels.add("过户单");
        boundaryLabels.add("购销合同");
        boundaryLabels.add("发货单位");
        boundaryLabels.add("提货单位");
        boundaryLabels.add("供方");
        boundaryLabels.add("需方");
        boundaryLabels.add("甲方");
        boundaryLabels.add("乙方");

        for (String label : labels) {
            String regex = Pattern.quote(label) + "\\s*[：:;；]\\s*(.+?)(?=(\\n|$|" +
                    boundaryLabels.stream().map(Pattern::quote).collect(Collectors.joining("|")) + "))";
            Matcher matcher = Pattern.compile(regex, Pattern.DOTALL).matcher(text);
            if (matcher.find()) {
                String value = cleanupValue(matcher.group(1));
                if (!isBlank(value)) {
                    return value;
                }
            }
        }
        return "";
    }

    private String extractValueLoose(String text, String... labels) {
        if (isBlank(text) || labels == null || labels.length == 0) return "";
        String[] lines = text.split("\\n");
        for (String rawLine : lines) {
            String line = rawLine == null ? "" : rawLine.trim();
            if (line.isEmpty()) continue;
            for (String label : labels) {
                if (!line.contains(label)) continue;
                String value = line;
                value = value.replaceFirst(".*" + Pattern.quote(label), "");
                value = value.replaceFirst("^[\\s:：;；\\-—_]+", "");
                value = value.replaceAll("\\s{2,}", " ").trim();
                value = value.replaceAll("^(是|为)\\s*", "").trim();
                value = cleanupValue(value);
                if (!isBlank(value) && value.length() >= 4) {
                    return value;
                }
            }
        }
        return "";
    }

    private List<String> extractCompanyCandidates(String text) {
        Set<String> result = new LinkedHashSet<>();
        Pattern p = Pattern.compile("([\\u4e00-\\u9fa5A-Za-z0-9（）()·\\-]{4,}?(?:有限责任公司|有限公司))");
        Matcher m = p.matcher(text);
        while (m.find()) {
            String name = cleanupValue(m.group(1));
            if (isBlank(name)) continue;
            result.add(name);
            if (result.size() >= 6) break;
        }
        return new ArrayList<>(result);
    }

    private String cleanupValue(String value) {
        if (value == null) return "";
        String v = value.replaceAll("[\\r\\n]+", " ")
                .replaceAll("\\s{2,}", " ")
                .replaceAll("^[：:;；\\s]+", "")
                .replaceAll("[：:;；\\s]+$", "")
                .trim();
        v = v.replaceAll("^(供方|需方|发货单位|提货单位|甲方|乙方)\\s*", "");
        v = v.replaceAll("^(合同编号|提单编号)\\s*", "");
        return v.trim();
    }

    private String cleanCompanyName(String raw) {
        String v = cleanupValue(raw);
        v = normalizeCommonOcrTypos(v);
        v = v.replaceAll("^(请过户至|请过户空交|请过户空至)\\s*", "");
        v = v.replaceAll("[^\\u4e00-\\u9fa5A-Za-z0-9（）()·\\-]", "");
        if (v.contains("请过户")) {
            return "";
        }
        v = v.replaceAll("(提货单位|提貨单位|发货单位|发貨单位).*", "");
        Matcher m = Pattern.compile("([\\u4e00-\\u9fa5A-Za-z0-9（）()·\\-]{4,}?(?:有限责任公司|有限公司))").matcher(v);
        if (m.find()) {
            return m.group(1);
        }
        if (!isLikelyCompanyName(v)) {
            return "";
        }
        return v;
    }

    private boolean isLikelyCompanyName(String s) {
        if (isBlank(s)) return false;
        String v = cleanupValue(s).replaceAll("\\s+", "");
        if (v.length() < 6) return false;
        if (v.matches("^[一二三四五六七八九十0-9\\-]+$")) return false;
        int zh = countChineseChars(v);
        return v.contains("公司") || zh >= 5;
    }

    private boolean isLikelyBillNo(String s) {
        if (isBlank(s)) return false;
        String v = normalizeBillNo(s);
        return v.matches("[A-Za-z0-9\\-_/]{8,}");
    }

    private String normalizeBillNo(String s) {
        if (isBlank(s)) return "";
        return cleanupValue(s).replaceAll("[^A-Za-z0-9\\-_/]", "");
    }

    private String extractCompanyAfterLabel(String text, String... labels) {
        if (isBlank(text) || labels == null) return "";
        for (String label : labels) {
            String regex = Pattern.quote(label) + "\\s*[：:;；]\\s*([\\u4e00-\\u9fa5A-Za-z0-9（）()·\\-]{3,}?(?:有限责任公司|有限公司))";
            Matcher m = Pattern.compile(regex).matcher(text);
            if (m.find()) {
                String c = cleanCompanyName(m.group(1));
                if (isLikelyCompanyName(c)) return c;
            }
        }
        return "";
    }

    private List<String> extractCompanyCandidatesForFallback(String text) {
        Set<String> result = new LinkedHashSet<>();
        if (isBlank(text)) return new ArrayList<>(result);
        String norm = normalizeCommonOcrTypos(text);
        Pattern p = Pattern.compile("([\\u4e00-\\u9fa5A-Za-z0-9（）()·\\-]{4,}?(?:有限责任公司|有限公司))");
        Matcher m = p.matcher(norm);
        while (m.find()) {
            String c = cleanCompanyName(m.group(1));
            if (isLikelyCompanyName(c)) {
                result.add(c);
            }
            if (result.size() >= 12) break;
        }
        return new ArrayList<>(result);
    }

    private String pickBestCompanyByInitials(List<String> candidates, String expected, String exclude) {
        if (candidates == null || candidates.isEmpty()) return "";
        String exp = isBlank(expected) ? "" : expected.toUpperCase(Locale.ROOT);
        String ex = isBlank(exclude) ? "" : cleanCompanyName(exclude);
        String best = "";
        int bestScore = Integer.MIN_VALUE;
        for (String c : candidates) {
            String cc = cleanCompanyName(c);
            if (!isLikelyCompanyName(cc)) continue;
            if (!isBlank(ex) && cc.equals(ex)) continue;
            int score = 0;
            String ini = toCompanyInitials(cc);
            if (!isBlank(exp) && !isBlank(ini)) {
                int n = Math.min(exp.length(), ini.length());
                for (int i = 0; i < n; i++) {
                    if (exp.charAt(i) == ini.charAt(i)) score += 3;
                    else break;
                }
            }
            // 发货单位一般更可能是“实业/化工”等而不是“物流”，给个轻微偏置
            if (cc.contains("实业")) score += 1;
            if (cc.contains("物流")) score -= 1;
            if (score > bestScore) {
                bestScore = score;
                best = cc;
            }
        }
        return best;
    }

    private String parseBillNoPart(String billNo, int partIndex) {
        if (isBlank(billNo) || partIndex < 1) return "";
        String[] arr = billNo.toUpperCase(Locale.ROOT).split("-");
        if (arr.length < partIndex) return "";
        return arr[partIndex - 1].replaceAll("[^A-Z0-9]", "");
    }

    private String repairBillNoByCompanies(String rawBillNo, String shipper, String consignee) {
        String bill = normalizeBillNo(rawBillNo).toUpperCase(Locale.ROOT);
        Matcher m = Pattern.compile("([A-Z0-9]{2,8})-([A-Z0-9]{2,8})-(\\d{8})-(\\d{3})").matcher(bill);
        if (!m.find()) {
            return bill;
        }
        String seg1 = m.group(1);
        String seg2 = m.group(2);
        String date = m.group(3);
        String seq = m.group(4);
        String expected1 = toCompanyInitials(shipper);
        String expected2 = toCompanyInitials(consignee);
        if (expected1.length() == 4) seg1 = expected1;
        if (expected2.length() == 4) seg2 = expected2;
        return seg1 + "-" + seg2 + "-" + date + "-" + seq;
    }

    private String toCompanyInitials(String company) {
        if (isBlank(company)) return "";
        String s = company.replaceAll("(有限责任公司|有限公司)$", "");
        HanyuPinyinOutputFormat format = new HanyuPinyinOutputFormat();
        format.setToneType(HanyuPinyinToneType.WITHOUT_TONE);
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < s.length(); i++) {
            char c = s.charAt(i);
            if (c >= '\u4e00' && c <= '\u9fa5') {
                try {
                    String[] py = PinyinHelper.toHanyuPinyinStringArray(c, format);
                    if (py != null && py.length > 0 && py[0].length() > 0) {
                        sb.append(Character.toUpperCase(py[0].charAt(0)));
                    }
                } catch (Exception ignored) {
                }
            } else if (Character.isLetter(c)) {
                sb.append(Character.toUpperCase(c));
            }
            if (sb.length() >= 4) break;
        }
        return sb.toString();
    }

    private String extractTransferPartyLineFallback(String text, boolean shipper) {
        if (isBlank(text)) return "";
        String[] lines = text.split("\\n");
        String[] keys = shipper
                ? new String[]{"发货单位", "发货单住", "发货单位", "发貨单位"}
                : new String[]{"提货单位", "提货单住", "提货单位", "提貨单位"};
        for (int i = 0; i < lines.length; i++) {
            String line = lines[i] == null ? "" : lines[i].trim();
            if (line.isEmpty()) continue;
            boolean hit = false;
            for (String k : keys) {
                if (line.contains(k)) {
                    hit = true;
                    break;
                }
            }
            if (!hit) continue;

            String value = line;
            int p = Math.max(value.indexOf(':'), Math.max(value.indexOf('：'), Math.max(value.indexOf(';'), value.indexOf('；'))));
            if (p >= 0 && p + 1 < value.length()) {
                value = value.substring(p + 1);
            }
            value = cleanCompanyName(value);
            if (isLikelyCompanyName(value)) {
                return value;
            }
            if (i + 1 < lines.length) {
                String next = cleanCompanyName(lines[i + 1]);
                if (isLikelyCompanyName(next)) {
                    return next;
                }
            }
        }
        return "";
    }

    private String normalizeCommonOcrTypos(String s) {
        if (s == null) return "";
        String v = s;
        v = v.replaceAll("有限公[司旬同可口回目日]", "有限公司");
        v = v.replaceAll("有限责任公[司旬同可口回目日]", "有限责任公司");
        v = v.replaceAll("公[旬同可口回目日]", "公司");
        return v;
    }

    private int countChineseChars(String s) {
        if (s == null) return 0;
        int c = 0;
        for (int i = 0; i < s.length(); i++) {
            char ch = s.charAt(i);
            if (ch >= '\u4e00' && ch <= '\u9fa5') c++;
        }
        return c;
    }

    private boolean isBlank(String s) {
        return s == null || s.trim().isEmpty();
    }

    private Path ensureUniqueTarget(Path source, String fileName) {
        Path parent = source.getParent();
        Path candidate = parent.resolve(fileName);
        if (!Files.exists(candidate) || candidate.equals(source)) {
            return candidate;
        }
        int dot = fileName.lastIndexOf('.');
        String base = dot >= 0 ? fileName.substring(0, dot) : fileName;
        String ext = dot >= 0 ? fileName.substring(dot) : "";
        int i = 1;
        while (true) {
            Path p = parent.resolve(base + "_" + i + ext);
            if (!Files.exists(p) || p.equals(source)) {
                return p;
            }
            i++;
        }
    }

    private String sanitizeFileName(String name) {
        if (name == null) return "";
        return name.replaceAll("[\\\\/:*?\"<>|]", "_")
                .replaceAll("\\s{2,}", " ")
                .trim();
    }

    private BufferedImage medianBlur3x3(BufferedImage src) {
        int w = src.getWidth();
        int h = src.getHeight();
        BufferedImage out = new BufferedImage(w, h, BufferedImage.TYPE_BYTE_GRAY);
        WritableRaster inR = src.getRaster();
        WritableRaster outR = out.getRaster();
        int[] window = new int[9];
        for (int y = 0; y < h; y++) {
            for (int x = 0; x < w; x++) {
                int idx = 0;
                for (int dy = -1; dy <= 1; dy++) {
                    for (int dx = -1; dx <= 1; dx++) {
                        int xx = clamp(x + dx, 0, w - 1);
                        int yy = clamp(y + dy, 0, h - 1);
                        window[idx++] = inR.getSample(xx, yy, 0);
                    }
                }
                java.util.Arrays.sort(window);
                outR.setSample(x, y, 0, window[4]);
            }
        }
        return out;
    }

    private BufferedImage otsuBinarize(BufferedImage src) {
        int w = src.getWidth();
        int h = src.getHeight();
        WritableRaster r = src.getRaster();
        int[] hist = new int[256];
        for (int y = 0; y < h; y++) {
            for (int x = 0; x < w; x++) {
                hist[r.getSample(x, y, 0)]++;
            }
        }
        int total = w * h;
        long sum = 0;
        for (int t = 0; t < 256; t++) sum += (long) t * hist[t];
        long sumB = 0;
        int wB = 0;
        double maxVar = -1;
        int threshold = 127;
        for (int t = 0; t < 256; t++) {
            wB += hist[t];
            if (wB == 0) continue;
            int wF = total - wB;
            if (wF == 0) break;
            sumB += (long) t * hist[t];
            double mB = (double) sumB / wB;
            double mF = (double) (sum - sumB) / wF;
            double varBetween = (double) wB * wF * (mB - mF) * (mB - mF);
            if (varBetween > maxVar) {
                maxVar = varBetween;
                threshold = t;
            }
        }
        BufferedImage out = new BufferedImage(w, h, BufferedImage.TYPE_BYTE_BINARY);
        WritableRaster outR = out.getRaster();
        for (int y = 0; y < h; y++) {
            for (int x = 0; x < w; x++) {
                int v = r.getSample(x, y, 0);
                outR.setSample(x, y, 0, v > threshold ? 1 : 0);
            }
        }
        return out;
    }

    private BufferedImage rotate(BufferedImage src, double angleDeg) {
        double rad = Math.toRadians(angleDeg);
        int w = src.getWidth();
        int h = src.getHeight();
        BufferedImage dst = new BufferedImage(w, h, src.getType());
        Graphics2D g2 = dst.createGraphics();
        try {
            g2.setRenderingHint(RenderingHints.KEY_INTERPOLATION, RenderingHints.VALUE_INTERPOLATION_BILINEAR);
            g2.setColor(Color.WHITE);
            g2.fillRect(0, 0, w, h);
            AffineTransform at = new AffineTransform();
            at.translate(w / 2.0, h / 2.0);
            at.rotate(rad);
            at.translate(-w / 2.0, -h / 2.0);
            g2.drawImage(src, at, null);
        } finally {
            g2.dispose();
        }
        return dst;
    }

    private int clamp(int v, int min, int max) {
        return Math.max(min, Math.min(max, v));
    }

    private static class TransferSlipFields {
        private final String shipper;
        private final String consignee;
        private final String billNo;

        private TransferSlipFields(String shipper, String consignee, String billNo) {
            this.shipper = shipper;
            this.consignee = consignee;
            this.billNo = billNo;
        }

        private static TransferSlipFields empty() {
            return new TransferSlipFields("", "", "");
        }
    }
}
