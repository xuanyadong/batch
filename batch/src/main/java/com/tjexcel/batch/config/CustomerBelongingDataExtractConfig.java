package com.tjexcel.batch.config;

import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.stereotype.Component;

/**
 * 客属表数据提取配置（功能7）。
 */
@Component
@ConfigurationProperties(prefix = "customer-belonging-data-extract")
public class CustomerBelongingDataExtractConfig {

    /** 客属表路径（按sheet读取企业抬头）。 */
    private String customerTablePath = "D:/test/batch/客属表.xlsx";

    /** 数据表路径（供方/需方匹配来源）。 */
    private String dataTablePath = "D:/test/batch/数据表.xlsx";

    /** 文件生成路径。 */
    private String outputDir = "D:/test/batch/output_customer_data_extract";

    /** 客属表公司列名。 */
    private String companyColumnName = "企业抬头";

    /** 需要输出的列，使用 | 分隔。 */
    private String extractColumns = "供方|需方|签约时间|产品名|规格型号|数量|单位|单价（含税）|价税合计金额|账面利润";


    public String getCustomerTablePath() {
        return customerTablePath;
    }

    public void setCustomerTablePath(String customerTablePath) {
        this.customerTablePath = customerTablePath;
    }

    public String getDataTablePath() {
        return dataTablePath;
    }

    public void setDataTablePath(String dataTablePath) {
        this.dataTablePath = dataTablePath;
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

    public String getExtractColumns() {
        return extractColumns;
    }

    public void setExtractColumns(String extractColumns) {
        this.extractColumns = extractColumns;
    }

}
