package com.tjexcel.batch.config;

import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.stereotype.Component;

/**
 * 客属表文件提取配置（功能6）。
 */
@Component
@ConfigurationProperties(prefix = "customer-belonging-extract")
public class CustomerBelongingExtractConfig {

    /** 客属数据表路径（包含多个sheet）。 */
    private String dataPath = "D:/test/batch/客属表.xlsx";

    /** 待扫描源压缩包路径。 */
    private String sourceZipPath = "D:/test/batch/source.zip";

    /** 输出目录（新压缩包与临时解压目录会在此目录下生成）。 */
    private String outputDir = "D:/test/batch/output_customer_extract";

    /** 企业抬头列名。 */
    private String companyColumnName = "企业抬头";

    public String getDataPath() {
        return dataPath;
    }

    public void setDataPath(String dataPath) {
        this.dataPath = dataPath;
    }

    public String getSourceZipPath() {
        return sourceZipPath;
    }

    public void setSourceZipPath(String sourceZipPath) {
        this.sourceZipPath = sourceZipPath;
    }

    public String getOutputDir() {
        return outputDir;
    }

    public void setOutputDir(String outputDir) {
        this.outputDir = outputDir;
    }

    public String getCompanyColumnName() {
        return companyColumnName;
    }

    public void setCompanyColumnName(String companyColumnName) {
        this.companyColumnName = companyColumnName;
    }
}
