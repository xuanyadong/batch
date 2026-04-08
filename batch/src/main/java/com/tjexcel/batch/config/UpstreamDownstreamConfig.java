package com.tjexcel.batch.config;

import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.stereotype.Component;

/**
 * Upstream/downstream table generation config.
 */
@Component
@ConfigurationProperties(prefix = "upstream-downstream")
public class UpstreamDownstreamConfig {

    /** Input data workbook path (same style as existing features). */
    private String dataPath = "D:/test/batch/数据模板.xlsx";

    /** Output directory. */
    private String outputDir = "D:/test/batch/output_upstream_downstream";

    /** Output file name pattern. Supports ${sheet}. */
    private String outputFileNamePattern = "上下游客户-${sheet}.xlsx";

    public String getDataPath() {
        return dataPath;
    }

    public void setDataPath(String dataPath) {
        this.dataPath = dataPath;
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
}
