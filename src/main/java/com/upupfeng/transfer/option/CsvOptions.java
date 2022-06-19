package com.upupfeng.transfer.option;

import java.io.Serializable;

/**
 * @author mawf
 */
public class CsvOptions implements Serializable {

    // 是否有表头
    private boolean withHeader = false;
    // 分隔符
    private String delimiter = ",";
    // 文件路径
    private String filePath;

    public String getDelimiter() {
        return delimiter;
    }

    public void setDelimiter(String delimiter) {
        this.delimiter = delimiter;
    }

    public String getFilePath() {
        return filePath;
    }

    public void setFilePath(String filePath) {
        this.filePath = filePath;
    }

    public boolean isWithHeader() {
        return withHeader;
    }

    public void setWithHeader(boolean withHeader) {
        this.withHeader = withHeader;
    }
}
