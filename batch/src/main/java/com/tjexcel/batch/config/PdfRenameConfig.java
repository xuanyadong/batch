package com.tjexcel.batch.config;

import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.stereotype.Component;

@Component
@ConfigurationProperties(prefix = "pdf-rename")
public class PdfRenameConfig {

    /** PDF scan directory. */
    private String scanDir = "D:/test/batch/pdf";

    /** Whether to scan recursively. */
    private boolean recursive = true;

    /** Tesseract executable absolute path. */
    private String tesseractExePath = "D:/Tesseract-OCR/tesseract.exe";

    /** OCR language. */
    private String language = "chi_sim";

    /** PDF render DPI. */
    private int dpi = 300;

    /** Max pages to OCR for each file, <=0 means all pages. */
    private int maxPages = 3;

    /** Dry-run mode: print rename plan only. */
    private boolean dryRun = true;

    /** Fast mode: skip multi-angle deskew retries for speed. */
    private boolean fastMode = true;

    /** Tesseract page segmentation mode. */
    private int psm = 6;

    public String getScanDir() {
        return scanDir;
    }

    public void setScanDir(String scanDir) {
        this.scanDir = scanDir;
    }

    public boolean isRecursive() {
        return recursive;
    }

    public void setRecursive(boolean recursive) {
        this.recursive = recursive;
    }

    public String getTesseractExePath() {
        return tesseractExePath;
    }

    public void setTesseractExePath(String tesseractExePath) {
        this.tesseractExePath = tesseractExePath;
    }

    public String getLanguage() {
        return language;
    }

    public void setLanguage(String language) {
        this.language = language;
    }

    public int getDpi() {
        return dpi;
    }

    public void setDpi(int dpi) {
        this.dpi = dpi;
    }

    public int getMaxPages() {
        return maxPages;
    }

    public void setMaxPages(int maxPages) {
        this.maxPages = maxPages;
    }

    public boolean isDryRun() {
        return dryRun;
    }

    public void setDryRun(boolean dryRun) {
        this.dryRun = dryRun;
    }

    public boolean isFastMode() {
        return fastMode;
    }

    public void setFastMode(boolean fastMode) {
        this.fastMode = fastMode;
    }

    public int getPsm() {
        return psm;
    }

    public void setPsm(int psm) {
        this.psm = psm;
    }
}
