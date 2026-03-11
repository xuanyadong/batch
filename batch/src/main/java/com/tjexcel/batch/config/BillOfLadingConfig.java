package com.tjexcel.batch.config;

import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.stereotype.Component;

/**
 * 提单批量生成配置
 */
@Component
@ConfigurationProperties(prefix = "bill-of-lading")
public class BillOfLadingConfig {

    /** 数据表路径 */
    private String dataPath = "D:/test/batch/提单数据.xlsx";

    /** 模板文件路径 */
    private String templatePath = "D:/test/batch/bill-of-lading.xlsx";

    /** 输出目录 */
    private String outputDir = "D:/test/batch/output_bill";

    /** 输出文件名模式，如：${供方简称}-${需方简称}-${提单编号}.xls */
    private String outputFileNamePattern = "${供方简称}-${需方简称}-${提单编号}.xls";

    /** 占位符前缀 */
    private String placeholderPrefix = "${";
    /** 占位符后缀 */
    private String placeholderSuffix = "}";

    /** 拆单：每份最小值（吨） */
    private int splitMin = 300;
    /** 拆单：每份最大值（吨） */
    private int splitMax = 495;
    /** 拆单：最大子单数量 */
    private int splitMaxSubCount = 40;

    public String getDataPath() {
        return dataPath;
    }

    public void setDataPath(String dataPath) {
        this.dataPath = dataPath;
    }

    public String getTemplatePath() {
        return templatePath;
    }

    public void setTemplatePath(String templatePath) {
        this.templatePath = templatePath;
    }

    public String getOutputDir() {
        return outputDir;
    }

    public void setOutputDir(String outputDir) {
        this.outputDir = outputDir;
    }

    public String getOutputFileNamePattern() {
        return outputFileNamePattern;
    }

    public void setOutputFileNamePattern(String outputFileNamePattern) {
        this.outputFileNamePattern = outputFileNamePattern;
    }

    public String getPlaceholderPrefix() {
        return placeholderPrefix;
    }

    public void setPlaceholderPrefix(String placeholderPrefix) {
        this.placeholderPrefix = placeholderPrefix;
    }

    public String getPlaceholderSuffix() {
        return placeholderSuffix;
    }

    public void setPlaceholderSuffix(String placeholderSuffix) {
        this.placeholderSuffix = placeholderSuffix;
    }

    public int getSplitMin() {
        return splitMin;
    }

    public void setSplitMin(int splitMin) {
        this.splitMin = splitMin;
    }

    public int getSplitMax() {
        return splitMax;
    }

    public void setSplitMax(int splitMax) {
        this.splitMax = splitMax;
    }

    public int getSplitMaxSubCount() {
        return splitMaxSubCount;
    }

    public void setSplitMaxSubCount(int splitMaxSubCount) {
        this.splitMaxSubCount = splitMaxSubCount;
    }
}
