package com.upupfeng.transfer.bean;

import java.io.Serializable;
import java.util.Iterator;

/**
 * @author mawf
 */
public class CsvData implements Serializable {

    // 迭代器
    Iterator<CsvLine> csvLineIterator;

    public CsvData(Iterator<CsvLine> csvLineIterator) {
        this.csvLineIterator = csvLineIterator;
    }

    public Iterator<CsvLine> getCsvLineIterator() {
        return csvLineIterator;
    }

    public void setCsvLineIterator(Iterator<CsvLine> csvLineIterator) {
        this.csvLineIterator = csvLineIterator;
    }

}
