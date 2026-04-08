package com.tjexcel.batch.config;

import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.stereotype.Component;

/**
 * Upstream/downstream append-write config.
 */
@Component
@ConfigurationProperties(prefix = "upstream-downstream")
public class UpstreamDownstreamConfig {

    /** Input data workbook path. */
    private String dataPath = "D:/test/batch/鏁版嵁妯℃澘.xlsx";

    /** Target upstream-downstream workbook path (append into this file). */
    private String targetFilePath = "D:/test/batch/鑱氫箼鐑笂涓嬫父瀹㈡埛260407.xlsx";

    /** Target sheet name in target workbook. Blank means first sheet. */
    private String targetSheetName = "";

    public String getDataPath() {
        return dataPath;
    }

    public void setDataPath(String dataPath) {
        this.dataPath = dataPath;
    }

    public String getTargetFilePath() {
        return targetFilePath;
    }

    public void setTargetFilePath(String targetFilePath) {
        this.targetFilePath = targetFilePath;
    }

    public String getTargetSheetName() {
        return targetSheetName;
    }

    public void setTargetSheetName(String targetSheetName) {
        this.targetSheetName = targetSheetName;
    }
}
