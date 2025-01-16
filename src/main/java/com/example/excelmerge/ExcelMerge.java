package com.example.excelmerge;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.read.metadata.ReadSheet;
import com.alibaba.excel.write.metadata.WriteSheet;

import java.util.*;

public class ExcelMerge {
    private static final String FILE_A = "/Users/zhymacbookair/IdeaProjects/excel-merge/a.xlsx";
    private static final String FILE_B = "/Users/zhymacbookair/IdeaProjects/excel-merge/b.xlsx";
    private static final String MAPPING_FILE = "/Users/zhymacbookair/IdeaProjects/excel-merge/mapping.xlsx";
    private static final String CONF_FILE = "/Users/zhymacbookair/IdeaProjects/excel-merge/conf.xlsx";
    private static final String OUTPUT_FILE = "/Users/zhymacbookair/IdeaProjects/excel-merge/merge.xlsx";

    private static final org.slf4j.Logger log = org.slf4j.LoggerFactory.getLogger(ExcelMerge.class);

    public static void main(String[] args) {
        ExcelMerge excelMerge = new ExcelMerge();
        excelMerge.mergeExcelFiles();
    }

    public void mergeExcelFiles() {
        // 读取配置文件
        log.info("开始读取配置文件...");
        ConfigInfo configInfo = readConfigFile();
        log.info("配置信息: sheet={}, keyColumn={}", configInfo.getSheetName(), configInfo.getKeyColumn());
        
        // 读取两个源文件的数据
        log.info("开始读取源文件...");
        List<Map<String, String>> dataA = readExcelFile(FILE_A);
        List<Map<String, String>> dataB = readExcelFile(FILE_B);
        log.info("文件A记录数: {}, 文件B记录数: {}", dataA.size(), dataB.size());
        
        // 获取所有的列名
        Set<String> allColumns = getAllColumns(dataA, dataB);
        log.info("合并后的所有列: {}", allColumns);
        
        // 读取映射数据
        log.info("开始读取映射文件...");
        Map<String, Map<String, String>> mappingData = readMappingFile(configInfo);
        log.info("映射数据记录数: {}", mappingData.size());
        
        // 合并数据并应用映射
        List<Map<String, String>> mergedData = mergeData(dataA, dataB, allColumns, mappingData, configInfo);
        log.info("合并后的记录数: {}", mergedData.size());
        
        // 写入结果文件
        log.info("开始写入结果文件...");
        writeToExcel(mergedData, allColumns);
        log.info("处理完成!");
    }

    private ConfigInfo readConfigFile() {
        List<Map<String, String>> configData = new ArrayList<>();
        EasyExcel.read(CONF_FILE, new AnalysisEventListener<Map<Integer, String>>() {
            @Override
            public void invoke(Map<Integer, String> data, AnalysisContext context) {
                Map<String, String> row = new HashMap<>();
                data.forEach((k, v) -> row.put(String.valueOf(k), v));
                configData.add(row);
            }

            @Override
            public void doAfterAllAnalysed(AnalysisContext context) {}
        }).sheet().doRead();

        ConfigInfo configInfo = new ConfigInfo();
        if (!configData.isEmpty()) {
            Map<String, String> config = configData.get(0);
            configInfo.setSheetName(config.get("0")); // 假设第一列是sheet名称
            configInfo.setKeyColumn(config.get("1")); // 假设第二列是关键列名称
        }
        return configInfo;
    }

    private List<Map<String, String>> readExcelFile(String filePath) {
        List<Map<String, String>> data = new ArrayList<>();
        EasyExcel.read(filePath, new AnalysisEventListener<Map<Integer, String>>() {
            @Override
            public void invoke(Map<Integer, String> rowData, AnalysisContext context) {
                Map<String, String> row = new HashMap<>();
                rowData.forEach((k, v) -> row.put(String.valueOf(k), v));
                data.add(row);
            }

            @Override
            public void doAfterAllAnalysed(AnalysisContext context) {}
        }).sheet().doRead();
        return data;
    }

    private Set<String> getAllColumns(List<Map<String, String>> dataA, List<Map<String, String>> dataB) {
        Set<String> columns = new HashSet<>();
        if (!dataA.isEmpty()) {
            columns.addAll(dataA.get(0).keySet());
        }
        if (!dataB.isEmpty()) {
            columns.addAll(dataB.get(0).keySet());
        }
        return columns;
    }

    private Map<String, Map<String, String>> readMappingFile(ConfigInfo configInfo) {
        Map<String, Map<String, String>> mappingData = new HashMap<>();
        List<Map<String, String>> data = new ArrayList<>();
        
        EasyExcel.read(MAPPING_FILE, new AnalysisEventListener<Map<Integer, String>>() {
            @Override
            public void invoke(Map<Integer, String> rowData, AnalysisContext context) {
                Map<String, String> row = new HashMap<>();
                rowData.forEach((k, v) -> row.put(String.valueOf(k), v));
                data.add(row);
            }

            @Override
            public void doAfterAllAnalysed(AnalysisContext context) {}
        }).sheet(configInfo.getSheetName()).doRead();

        // 构建映射关系
        for (Map<String, String> row : data) {
            String key = row.get(configInfo.getKeyColumn());
            if (key != null) {
                mappingData.put(key, row);
            }
        }
        
        return mappingData;
    }

    private List<Map<String, String>> mergeData(
            List<Map<String, String>> dataA,
            List<Map<String, String>> dataB,
            Set<String> allColumns,
            Map<String, Map<String, String>> mappingData,
            ConfigInfo configInfo) {
        
        List<Map<String, String>> mergedData = new ArrayList<>();
        mergedData.addAll(dataA);
        mergedData.addAll(dataB);

        // 应用映射数据
        for (Map<String, String> row : mergedData) {
            String key = row.get(configInfo.getKeyColumn());
            if (key != null && mappingData.containsKey(key)) {
                Map<String, String> mappingRow = mappingData.get(key);
                row.putAll(mappingRow);
            }
        }

        return mergedData;
    }

    private void writeToExcel(List<Map<String, String>> mergedData, Set<String> columns) {
        List<List<String>> rows = new ArrayList<>();
        
        // 添加表头
        List<String> header = new ArrayList<>(columns);
        rows.add(header);
        
        // 添加数据行
        for (Map<String, String> row : mergedData) {
            List<String> dataRow = new ArrayList<>();
            for (String column : columns) {
                dataRow.add(row.getOrDefault(column, ""));
            }
            rows.add(dataRow);
        }
        
        // 写入Excel文件
        EasyExcel.write(OUTPUT_FILE).sheet("合并结果").doWrite(rows);
    }

    private static class ConfigInfo {
        private String sheetName;
        private String keyColumn;

        public String getSheetName() {
            return sheetName;
        }

        public void setSheetName(String sheetName) {
            this.sheetName = sheetName;
        }

        public String getKeyColumn() {
            return keyColumn;
        }

        public void setKeyColumn(String keyColumn) {
            this.keyColumn = keyColumn;
        }
    }
}
