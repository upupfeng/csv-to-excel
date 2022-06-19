package com.upupfeng.transfer;

import com.upupfeng.transfer.option.CsvOptions;
import com.upupfeng.transfer.option.OutputOptions;

/**
 * @author mawf
 */
public class App {

    public static void main(String[] args) throws Exception {

        CsvOptions csvOptions = new CsvOptions();
        // 分隔符
        csvOptions.setDelimiter("|");
        // csv文件路径
        csvOptions.setFilePath("D:\\dev\\csv\\a.csv");
        csvOptions.setWithHeader(false);

        OutputOptions outputOptions = new OutputOptions();
        // 结果文件目录
        outputOptions.setFileDir("D:\\dev\\csv\\1");
        // 结果文件名称
        outputOptions.setFileName("测试.xlsx");

        CsvToExcelTransfer csvTransfer = new CsvToExcelTransfer();
        csvTransfer.csvToExcel(csvOptions, outputOptions);

    }


}
