package com.tjexcel.batch.config;

import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.stereotype.Component;

import java.util.ArrayList;
import java.util.List;

/**
 * 资金流批量生成配置
 * 生成一个 Excel，表现资金流向：付款方(需方) → 收款方(供方) 金额
 */
@Component
@ConfigurationProperties(prefix = "fund-flow")
public class FundFlowConfig {

    /** 数据表路径（可与合同/提单共用） */
    private String dataPath = "D:/test/batch/资金流数据.xlsx";

    /** 输出目录 */
    private String outputDir = "D:/test/batch/output_fund";

    /** 输出文件名（一个 Excel 文件，如：资金流.xlsx） */
    private String outputFileName = "资金流.xlsx";

    /** 金额列名，默认价税合计金额 */
    private String amountColumn = "价税合计金额";

    /** 资金流起点公司全称列表（逗号分隔），按顺序逐个展开 */
    private List<String> rootCompanies = new ArrayList<>();

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

    public String getOutputFileName() {
        return outputFileName;
    }

    public void setOutputFileName(String outputFileName) {
        this.outputFileName = outputFileName;
    }

    public String getAmountColumn() {
        return amountColumn;
    }

    public void setAmountColumn(String amountColumn) {
        this.amountColumn = amountColumn;
    }

    public List<String> getRootCompanies() {
        return rootCompanies;
    }

    public void setRootCompanies(List<String> rootCompanies) {
        this.rootCompanies = rootCompanies;
    }
}
