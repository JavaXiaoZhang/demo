package com.example.demo;

import javafx.fxml.FXML;
import javafx.scene.control.Alert;
import javafx.stage.DirectoryChooser;
import javafx.stage.Window;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.text.NumberFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class HelloController {
    static final String compareFile = "D:\\compare.xls";
    static final String[] titles = {"报表代码", "指标代码", "本期数据值", "上期数据值", "比上期（万元）", "比上期（%）", "备注"};
    Map<String, List<String>> compareMap = new HashMap<>();
    Map<String, List<Double>> previousMap = new HashMap<>();
    Map<String, List<Double>> currentMap = new HashMap<>();

    @FXML
    protected void onPreviousButtonClick() {
        DirectoryChooser chooser = new DirectoryChooser();
        //设置文件上传类型
        //FileChooser.ExtensionFilter extFilter = new FileChooser.ExtensionFilter("excel files (*.xlsx)", "*.xlsx");
        //chooser.getExtensionFilters().add(extFilter);
        //chooser.setInitialDirectory(new File("E:\\workspace\\demo"));
        chooser.setTitle("上传上期文件夹");
        File file = chooser.showDialog(Window.impl_getWindows().next());
        if (file != null) {
            try {
                // 解析文件夹
                doHandlerDirectory(file, previousMap);
            } catch (Exception e) {
                Alert alert = new Alert(Alert.AlertType.WARNING);
                alert.titleProperty().set("警告");
                alert.headerTextProperty().set(e.getMessage());
                alert.showAndWait();
            }
        }
    }

    @FXML
    protected void onCurrentButtonClick() {
        DirectoryChooser chooser = new DirectoryChooser();
        //chooser.setInitialDirectory(new File("E:\\workspace\\demo"));
        chooser.setTitle("上传当期文件夹");
        File file = chooser.showDialog(Window.impl_getWindows().next());
        if (file != null) {
            try {
                // 解析文件夹
                doHandlerDirectory(file, currentMap);
            } catch (Exception e) {
                Alert alert = new Alert(Alert.AlertType.WARNING);
                alert.titleProperty().set("警告");
                alert.headerTextProperty().set(e.getMessage());
                alert.showAndWait();
            }
        }
    }

    private void doHandlerDirectory(File file, Map<String, List<Double>> valueMap) throws Exception {
        // 解析比对文件
        doHandlerCompareFile();

        // 将上次内容清空
        valueMap.clear();

        for (File listFile : file.listFiles()) {
            Workbook workbook = new HSSFWorkbook(new FileInputStream(listFile));
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                Sheet sheet = workbook.getSheetAt(i);
                // 报表代码
                String sheetCode = sheet.getSheetName();
                // 若比较文档中不包含该报表代码，则不用处理
                List<String> indexList = compareMap.get(sheetCode);
                if (indexList == null || indexList.isEmpty()) {
                    continue;
                }

                List<Double> valueList = new ArrayList<>();
                valueMap.put(sheetCode, valueList);

                for (String index : indexList) {
                    char ch = index.charAt(0);
                    int rowIndex = Integer.valueOf(index.substring(1)) - 1;
                    int cellIndex = ch - 'A';
                    valueList.add(sheet.getRow(rowIndex).getCell(cellIndex).getNumericCellValue());
                }
            }
        }

        // 如果当期不为空，上期为空则报错提示
        if (!currentMap.isEmpty() && previousMap.isEmpty()) {
            throw new RuntimeException("请上传当期文件夹");
        }

        if (!currentMap.isEmpty() && !previousMap.isEmpty()) {
            // 开始比对
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sheet = wb.createSheet("数据比对结果");
            HSSFRow firstRow = sheet.createRow(0);
            for (int i = 0; i < titles.length; i++) {
                HSSFCell cell = firstRow.createCell(i);
                cell.setCellValue(titles[i]);
            }

            int i = 1;
            for (Map.Entry<String, List<String>> entry : compareMap.entrySet()) {
                String sheetCode = entry.getKey();
                List<Double> previousValueList = previousMap.get(sheetCode);
                List<Double> currentValueList = currentMap.get(sheetCode);
                for (int j = 0; j < entry.getValue().size(); j++) {
                    Double previousValue = previousValueList.get(j);
                    Double currentValue = currentValueList.get(j);
                    /**
                     * 当期和前期进行比较
                     * 比上期变动幅度大于100% 棕黄色
                     * 比上期变动幅度等于100% 黄色
                     * 比上期变动幅度小于-100% 红色
                     * 本期数据为0，上期不等于0 绿色
                     * 本期数据不等于0，上期数据为0 蓝色
                     */
                    String remarks;
                    CellStyle cellStyle = wb.createCellStyle();
                    cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                    if (previousValue != 0 && previousValue * 2 < currentValue) {
                        cellStyle.setFillForegroundColor(HSSFColor.HSSFColorPredefined.ORANGE.getIndex());
                        remarks = "比上期变动幅度大于100%";
                        //logger.info("{}设置棕黄色", index);
                    } else if (previousValue != 0 && previousValue * 2 == currentValue) {
                        cellStyle.setFillForegroundColor(HSSFColor.HSSFColorPredefined.YELLOW.getIndex());
                        remarks = "比上期变动幅度等于100%";
                        //logger.info("{}设置黄色", index);
                    } else if (currentValue != 0 && previousValue > currentValue * 2) {
                        cellStyle.setFillForegroundColor(HSSFColor.HSSFColorPredefined.RED.getIndex());
                        remarks = "比上期变动幅度小于-100%";
                        //logger.info("{}设置红色", index);
                    } else if (previousValue != 0 && currentValue == 0) {
                        cellStyle.setFillForegroundColor(HSSFColor.HSSFColorPredefined.GREEN.getIndex());
                        remarks = "本期数据为0，上期不等于0";
                        //logger.info("{}设置绿色", index);
                    } else if (previousValue == 0 && currentValue != 0) {
                        cellStyle.setFillForegroundColor(HSSFColor.HSSFColorPredefined.BLUE.getIndex());
                        remarks = "本期数据不等于0，上期数据为0";
                        //logger.info("{}设置蓝色", index);
                    } else {
                        continue;
                    }

                    HSSFRow row = sheet.createRow(i++);
                    HSSFCell cell0 = row.createCell(0);
                    cell0.setCellValue(sheetCode);
                    HSSFCell cell1 = row.createCell(1);
                    cell1.setCellValue(entry.getValue().get(j));
                    HSSFCell cell2 = row.createCell(2, CellType.NUMERIC);
                    cell2.setCellValue(currentValue);
                    HSSFCell cell3 = row.createCell(3, CellType.NUMERIC);
                    cell3.setCellValue(previousValue);
                    HSSFCell cell4 = row.createCell(4);
                    cell4.setCellValue((currentValue - previousValue) / 10000);
                    HSSFCell cell5 = row.createCell(5);
                    NumberFormat numberFormat = NumberFormat.getInstance();
                    numberFormat.setMinimumFractionDigits(2);
                    cell5.setCellValue(numberFormat.format(currentValue / previousValue) + "%");
                    HSSFCell cell6 = row.createCell(6);
                    cell6.setCellValue(remarks);

                    for (Cell cell : row) {
                        cell.setCellStyle(cellStyle);
                    }
                }
            }

            File resultFile = new File("D:\\数据比对结果.xls");
            wb.write(resultFile);

            Alert alert = new Alert(Alert.AlertType.INFORMATION);
            alert.titleProperty().set("解析完成");
            alert.headerTextProperty().set("已生成结果文件 D:\\数据比对结果.xls");
            alert.showAndWait();
        }
    }

    private void doHandlerCompareFile() throws Exception {
        File file = new File(compareFile);
        if (!file.exists()) {
            throw new RuntimeException("请把比对文件放在指定位置 D:\\compare.xls");
        }

        // 将上次内容清空
        compareMap.clear();

        Workbook compareWorkbook = new HSSFWorkbook(new FileInputStream(file));
        Sheet compareWorkbookSheet = compareWorkbook.getSheetAt(0);
        for (int i = 0; i <= compareWorkbookSheet.getLastRowNum(); i++) {
            Row row = compareWorkbookSheet.getRow(i);
            String sheetCode = row.getCell(0).getStringCellValue().trim();
            String indexCode = row.getCell(1).getStringCellValue().trim();
            List<String> indexCodeList;
            if (compareMap.containsKey(sheetCode)) {
                indexCodeList = compareMap.get(sheetCode);
            } else {
                indexCodeList = new ArrayList<>();
                compareMap.put(sheetCode, indexCodeList);
            }
            indexCodeList.add(indexCode);
        }
    }
}