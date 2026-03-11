package com.tjexcel.batch.config;

import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.stereotype.Component;

/**
 * 合同批量生成配置
 */
@Component
@ConfigurationProperties(prefix = "contract")
public class ContractConfig {

    /**
     * 数据表路径（如 0126.xlsx），第一个 sheet 为数据源
     */
    private String dataPath = "D:/test/batch/0126.xlsx";

    /**
     * 模板文件路径
     */
    private String templatePath = "D:/test/batch/muban.xlsx";

    /**
     * 生成文件的输出目录
     */
    private String outputDir = "D:/test/batch/output";

    /**
     * 输出文件名模式，使用 ${列名} 占位符，如：合同-${买方}-${卖方}-${合同编号}.xls
     */
    private String outputFileNamePattern = "合同-${买方}-${卖方}-${合同编号}.xls";

    /**
     * 模板中的占位符格式，默认 ${列名}
     */
    private String placeholderPrefix = "${";
    private String placeholderSuffix = "}";

    /**
     * 列-单元格映射：数据列的值写入模板指定单元格。格式 "B3=A,D5=B,E7=C"
     */
    private String columnCellMapping = "";

    /**
     * 已填充样本路径，用于自动推断列-单元格映射
     */
    private String sampleFilledPath = "";

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

    public String getColumnCellMapping() {
        return columnCellMapping;
    }

    public void setColumnCellMapping(String columnCellMapping) {
        this.columnCellMapping = columnCellMapping;
    }

    public String getSampleFilledPath() {
        return sampleFilledPath;
    }

    public void setSampleFilledPath(String sampleFilledPath) {
        this.sampleFilledPath = sampleFilledPath;
    }
}
