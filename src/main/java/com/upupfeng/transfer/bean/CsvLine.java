package com.upupfeng.transfer.bean;

import java.util.Arrays;
import java.util.stream.Collectors;

/**
 * @author mawf
 */
public class CsvLine {

    private boolean headerLine = false;

    private String[] valueArr;

    public CsvLine(String[] valueArr) {
        this.valueArr = valueArr;
    }

    public String[] getValueArr() {
        return valueArr;
    }

    public void setValueArr(String[] valueArr) {
        this.valueArr = valueArr;
    }

    public boolean isHeaderLine() {
        return headerLine;
    }

    public void setHeaderLine(boolean headerLine) {
        this.headerLine = headerLine;
    }

    @Override
    public String toString() {
        return Arrays.stream(valueArr).collect(Collectors.joining(","));
    }

}
