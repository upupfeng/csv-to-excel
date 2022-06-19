package com.upupfeng.transfer;

import com.upupfeng.transfer.bean.CsvData;
import com.upupfeng.transfer.bean.CsvLine;
import com.upupfeng.transfer.option.CsvOptions;
import com.upupfeng.transfer.option.OutputOptions;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.*;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.regex.Pattern;

/**
 * https://www.jb51.net/article/157349.htm
 *
 * @author mawf
 */
public class CsvToExcelTransfer {

    // csv转为excel
    public void csvToExcel(CsvOptions csvOptions, OutputOptions outputOptions) throws Exception {
        System.out.println("开始执行将csv转成excel");
        CsvData csvData = readCsv(csvOptions);
        writeCsvDataToExcel(outputOptions, csvOptions, csvData);
        System.out.println("将csv转成excel完成");
    }

    // 读取csv
    private CsvData readCsv(CsvOptions csvOptions) throws Exception {
        String delimiter = csvOptions.getDelimiter();
        String filePath = csvOptions.getFilePath();
        BufferedReader reader;
        try {
            reader = new BufferedReader(new FileReader(filePath));
            BufferedReader finalReader = reader;
            Iterator<CsvLine> iterator = new Iterator<CsvLine>() {
                private String line = null;
                private int lineNum = 0;

                @Override
                public boolean hasNext() {
                    try {
                        return (line = finalReader.readLine()) != null;
                    } catch (IOException e) {
                        e.printStackTrace();
                        return false;
                    }
                }

                @Override
                public CsvLine next() {
                    lineNum++;
                    String[] valueArr = line.split(Pattern.quote(delimiter));
                    CsvLine csvLine = new CsvLine(valueArr);
                    if (lineNum == 1) {
                        if (csvOptions.isWithHeader()) {
                            csvLine.setHeaderLine(true);
                        }
                    }
                    return csvLine;
                }
            };

            return new CsvData(iterator);
        } catch (Exception e) {
            throw e;
        } finally {
            // TODO 关闭会导致Stream closed的错
//            if (reader != null) {
//                try {
//                    reader.close();
//                } catch (Exception e) {
//                }
//            }
        }
    }

    // 将csv数据写入excel
    private void writeCsvDataToExcel(OutputOptions outputOptions, CsvOptions csvOptions,
                                     CsvData csvData) {
        // 将数据填充到workbook中
        Workbook workbook = fillDataToWorkbook(csvData);
        String resFileFullPath = buildResFileFullPath(outputOptions, csvOptions);
        writeWorkbookToFile(workbook, resFileFullPath);
    }

    // 构建结果文件的全路径
    private String buildResFileFullPath(OutputOptions outputOptions, CsvOptions csvOptions) {
        String fileDir = outputOptions.getFileDir();
        if (fileDir == null) {
            File file = new File(csvOptions.getFilePath());
            fileDir = file.getParent();
        }
        String fileName = outputOptions.getFileName();
        if (fileName == null) {
            File file = new File(csvOptions.getFilePath());
            String name = file.getName();
            fileName = name.substring(0, name.lastIndexOf(".")) + ".xlsx";
        }
        // 结果文件全路径
        String resFileFullPath = new File(fileDir, fileName).getPath();
        return resFileFullPath;
    }

    //将workbook写入文件
    private void writeWorkbookToFile(Workbook workbook, String resFileFullPath) {
        FileOutputStream os = null;
        try {
            File file = new File(resFileFullPath);
            File parentFile = file.getParentFile();
            if (!parentFile.exists()) {
                parentFile.mkdirs();
            }
            if (!file.exists()) {
                file.createNewFile();
            }
            os = new FileOutputStream(file);
            workbook.write(os);
            os.flush();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (os != null) {
                try {
                    os.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }

    // 创建sheet
    private Sheet createSheet(Workbook workbook) {
        Sheet sheet = workbook.createSheet();
        sheet.setDefaultColumnWidth(20);
        sheet.setDefaultRowHeight((short) 300);
        return sheet;
    }

    // 将数据填充到workbook中
    private Workbook fillDataToWorkbook(CsvData csvData) {
        // 获取迭代器
        Iterator<CsvLine> csvLineIterator = csvData.getCsvLineIterator();

        // 生成xlsx的Excel
        Workbook workbook = new SXSSFWorkbook();
        Sheet sheet = createSheet(workbook);

        // 列名到内容最大长度的映射。
        Map<Integer, Integer> columnIndexToLengthMap = new HashMap<>();

        // 行数，从0开始
        int rowNum = 0;
        while (csvLineIterator.hasNext()) {
            // 创建行
            Row row = sheet.createRow(rowNum);
            rowNum++;

            CsvLine csvLine = csvLineIterator.next();
            String[] valueArr = csvLine.getValueArr();
            // 获取每个字段
            for (int i = 0; i < valueArr.length; i++) {
                // 创建每一个单元格
                Cell cell = row.createCell(i);
                cell.setCellValue(valueArr[i]);
                if (csvLine.isHeaderLine()) {
                    cell.setCellStyle(buildHeaderCellStyle(workbook));
                } else {
                    cell.setCellStyle(buildTextCellStyle(workbook));
                }
                // 获取最大列宽
                int length = Math.min(cell.getStringCellValue().getBytes().length * 256 + 512, 12000);
                if (columnIndexToLengthMap.get(i) == null || columnIndexToLengthMap.get(i) < length) {
                    columnIndexToLengthMap.put(i, length);
                }
            }
        }

        // 自适应列宽
        for (int i : columnIndexToLengthMap.keySet()) {
            sheet.setColumnWidth(i, columnIndexToLengthMap.get(i));
        }

        return workbook;
    }

    // 构建正文的
    private CellStyle buildTextCellStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        //对齐方式设置
        style.setAlignment(HorizontalAlignment.CENTER);
        //边框颜色和宽度设置
        style.setBorderBottom(BorderStyle.THIN);
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex()); // 下边框
        style.setBorderLeft(BorderStyle.THIN);
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex()); // 左边框
        style.setBorderRight(BorderStyle.THIN);
        style.setRightBorderColor(IndexedColors.BLACK.getIndex()); // 右边框
        style.setBorderTop(BorderStyle.THIN);
        style.setTopBorderColor(IndexedColors.BLACK.getIndex()); // 上边框
        //粗体字设置
        Font font = workbook.createFont();
        font.setBold(false);
        style.setFont(font);
        return style;
    }

    // 构建表头的
    private CellStyle buildHeaderCellStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        //对齐方式设置
        style.setAlignment(HorizontalAlignment.CENTER);
        //边框颜色和宽度设置
        style.setBorderBottom(BorderStyle.THIN);
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex()); // 下边框
        style.setBorderLeft(BorderStyle.THIN);
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex()); // 左边框
        style.setBorderRight(BorderStyle.THIN);
        style.setRightBorderColor(IndexedColors.BLACK.getIndex()); // 右边框
        style.setBorderTop(BorderStyle.THIN);
        style.setTopBorderColor(IndexedColors.BLACK.getIndex()); // 上边框
        //设置背景颜色
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        //粗体字设置
        Font font = workbook.createFont();
        font.setBold(false);
        style.setFont(font);
        return style;
    }

    public static void main(String[] args) throws Exception {

        CsvOptions csvOptions = new CsvOptions();
        csvOptions.setDelimiter("|");
        csvOptions.setFilePath("D:\\dev\\csv\\a.csv");
        csvOptions.setWithHeader(true);

        OutputOptions outputOptions = new OutputOptions();
        outputOptions.setFileDir("D:\\dev\\csv");
        outputOptions.setFileName("测试.xlsx");

        CsvToExcelTransfer csvTransfer = new CsvToExcelTransfer();
        csvTransfer.csvToExcel(csvOptions, outputOptions);

        // 测试转csv
//        CsvToExcelTransfer csvTransfer = new CsvToExcelTransfer();
//        CsvData csvData = csvTransfer.readCsv(csvOptions);
//        Iterator<CsvLine> csvLineIterator = csvData.getCsvLineIterator();
//        while (csvLineIterator.hasNext()) {
//            CsvLine next = csvLineIterator.next();
//
//            System.out.println(next);
//        }

    }

    private static void testCsvTransfer() throws Exception {

        CsvOptions csvOptions = new CsvOptions();
        csvOptions.setDelimiter("|");
        csvOptions.setFilePath("D:\\dev\\csv\\a.csv");

        OutputOptions outputOptions = new OutputOptions();
        outputOptions.setFileDir("D:\\dev\\csv");
        outputOptions.setFileName("测试.xlsx");

        // 测试转csv
        CsvToExcelTransfer csvTransfer = new CsvToExcelTransfer();
        CsvData csvData = csvTransfer.readCsv(csvOptions);
        Iterator<CsvLine> csvLineIterator = csvData.getCsvLineIterator();
        while (csvLineIterator.hasNext()) {
            CsvLine next = csvLineIterator.next();

            System.out.println(next);
        }

    }

}
