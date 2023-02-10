package com.example.demo;

import javafx.fxml.FXML;
import javafx.fxml.FXMLLoader;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.scene.control.Alert;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.stage.DirectoryChooser;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import javafx.stage.Window;
import org.apache.commons.collections4.multimap.ArrayListValuedHashMap;
import org.apache.commons.io.FilenameUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.NumberFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class HelloController {
    static final String USER_DIR = System.getProperty("user.dir");
    static final String COMPARE_FILE_PATH = USER_DIR + File.separator + "compare.xls";
    static final String[] titles = {"报表代码", "指标位置", "指标代码", "指标名称", "本期数据值", "上期数据值", "比上期（万元）", "比上期（%）", "备注"};
    static Map<String, List<String>> compareMap = new HashMap<>();
    static ArrayListValuedHashMap<String, String> locationMap = new ArrayListValuedHashMap<>();
    static Map<String, List<Double>> previousMap = new HashMap<>();
    static Map<String, List<Double>> currentMap = new HashMap<>();

    @FXML
    Button compareButton;

    @FXML
    protected void onUploadButtonClick() {
        FileChooser chooser = new FileChooser();
        //设置文件上传类型
        FileChooser.ExtensionFilter extFilter = new FileChooser.ExtensionFilter("excel files (*.xls)", "*.xls", "*.xlsx");
        chooser.getExtensionFilters().add(extFilter);
        chooser.setTitle("上传校验文件");
        File file = chooser.showOpenDialog(Window.impl_getWindows().next());
        //File file = chooser.showOpenDialog(Window.getWindows().get(1));
        try {
            doHandlerCompareFile(file);
        } catch (Exception e) {
            alert(Alert.AlertType.WARNING, "警告", e.getMessage());
        }
    }

    @FXML
    protected void onDownloadButtonClick() {
        try {
            File file = new File(COMPARE_FILE_PATH);
            if (compareMap.isEmpty() && !file.exists()) {
                throw new RuntimeException("未检测到校验文件");
            }
            if (!compareMap.isEmpty()) {
                HSSFWorkbook workbook = new HSSFWorkbook();
                HSSFSheet sheet = workbook.createSheet();
                int i = 0;
                for (Map.Entry<String, List<String>> entry : compareMap.entrySet()) {
                    String sheetCode = entry.getKey();
                    for (int j = 0; j < entry.getValue().size(); j++) {
                        String value = entry.getValue().get(j);
                        HSSFRow row = sheet.createRow(i++);
                        int k = 0;
                        HSSFCell cell = row.createCell(k++);
                        cell.setCellValue(sheetCode);
                        HSSFCell cell1 = row.createCell(k++);
                        cell1.setCellValue(value);
                        HSSFCell cell2 = row.createCell(k++);
                        List<String> strings = locationMap.get(value);
                        cell2.setCellValue(strings.get(0));
                        HSSFCell cell3 = row.createCell(k++);
                        cell3.setCellValue(strings.get(1));
                    }
                }
                workbook.write(file);
            }
            alert(Alert.AlertType.INFORMATION, "下载完成", "校验文件已下载至" + COMPARE_FILE_PATH);
        } catch (Exception e) {
            alert(Alert.AlertType.WARNING, "警告", e.getMessage());
        }
    }

    void alert(Alert.AlertType alertType, String title, String text) {
        Alert alert = new Alert(alertType);
        alert.titleProperty().set(title);
        alert.headerTextProperty().set(text);
        alert.showAndWait();
    }

    @FXML
    protected void onStartButtonClick() {
        try {
            if (previousMap.isEmpty()) {
                throw new RuntimeException("请上传上期文件夹");
            }
            if (currentMap.isEmpty()) {
                throw new RuntimeException("请上传当期文件夹");
            }
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
                        cellStyle.setFillForegroundColor(HSSFColor.HSSFColorPredefined.LIGHT_BLUE.getIndex());
                        remarks = "本期数据不等于0，上期数据为0";
                        //logger.info("{}设置蓝色", index);
                    } else {
                        continue;
                    }

                    HSSFRow row = sheet.createRow(i++);
                    int columnIndex = 0;
                    HSSFCell cell0 = row.createCell(columnIndex++);
                    cell0.setCellValue(sheetCode);
                    HSSFCell cell1 = row.createCell(columnIndex++);
                    String indexLocation = entry.getValue().get(j);
                    cell1.setCellValue(indexLocation);
                    HSSFCell cellIndexCode = row.createCell(columnIndex++);
                    List<String> strings = locationMap.get(indexLocation);
                    cellIndexCode.setCellValue(strings.get(0));
                    HSSFCell cellIndexName = row.createCell(columnIndex++);
                    cellIndexName.setCellValue(strings.get(1));
                    HSSFCell cell2 = row.createCell(columnIndex++, CellType.NUMERIC);
                    cell2.setCellValue(currentValue);
                    HSSFCell cell3 = row.createCell(columnIndex++, CellType.NUMERIC);
                    cell3.setCellValue(previousValue);
                    HSSFCell cell4 = row.createCell(columnIndex++);
                    cell4.setCellValue((currentValue - previousValue) / 10000);
                    HSSFCell cell5 = row.createCell(columnIndex++);
                    NumberFormat numberFormat = NumberFormat.getInstance();
                    numberFormat.setMinimumFractionDigits(2);
                    cell5.setCellValue(numberFormat.format(currentValue / previousValue) + "%");
                    HSSFCell cell6 = row.createCell(columnIndex++);
                    cell6.setCellValue(remarks);

                    for (Cell cell : row) {
                        cell.setCellStyle(cellStyle);
                    }
                }
            }

            String resultFileName = USER_DIR + File.separator + "数据比对结果.xls";
            File resultFile = new File(resultFileName);
            wb.write(resultFile);

            Alert alert = new Alert(Alert.AlertType.INFORMATION);
            alert.titleProperty().set("解析完成");
            alert.headerTextProperty().set("已生成结果文件 " + resultFileName);
            alert.showAndWait();
        } catch (Exception e) {
            alert(Alert.AlertType.WARNING, "警告", e.getMessage());
        }
    }

    @FXML
    protected void onCompareButtonClick() throws IOException {
        // 创建新的stage
        Stage secondStage = new Stage();
        FXMLLoader fxmlLoader = new FXMLLoader(HelloApplication.class.getResource("second-view.fxml"));
        Parent root = fxmlLoader.load();
        Label node = (Label) root.getChildrenUnmodifiable().get(0);
        node.setText("校验文件默认位置 " + COMPARE_FILE_PATH);
        Scene secondScene = new Scene(root, 320, 240);

        secondStage.setTitle("校验文件");
        secondStage.setScene(secondScene);
        secondStage.show();
    }

    @FXML
    protected void onPreviousButtonClick() {
        DirectoryChooser chooser = new DirectoryChooser();
        chooser.setTitle("上传上期文件夹");
        File file = chooser.showDialog(Window.impl_getWindows().next());
        //File file = chooser.showDialog(Window.getWindows().get(0));
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
        //File file = chooser.showDialog(Window.getWindows().get(0));
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
        doHandlerCompareFile(null);

        // 将上次内容清空
        valueMap.clear();

        for (File listFile : file.listFiles()) {
            String sheetCode = FilenameUtils.getBaseName(listFile.getName());
            // 若比较文档中不包含该报表代码，则不用处理
            List<String> indexList = compareMap.get(sheetCode);
            if (indexList == null || indexList.isEmpty()) {
                continue;
            }
            Workbook workbook = new HSSFWorkbook(new FileInputStream(listFile));
            Sheet sheet = workbook.getSheetAt(0);
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

    private void doHandlerCompareFile(File file) throws Exception {
        if (file == null && !compareMap.isEmpty()) {
            return;
        }
        if (file == null) {
            file = new File(COMPARE_FILE_PATH);
        }

        if (file == null || !file.exists()) {
            throw new RuntimeException("请手动上传校验文件或将校验文件放在指定位置 " + COMPARE_FILE_PATH);
        }

        // 将上次内容清空
        compareMap.clear();
        locationMap.clear();

        Workbook compareWorkbook = new HSSFWorkbook(new FileInputStream(file));
        Sheet compareWorkbookSheet = compareWorkbook.getSheetAt(0);
        for (int i = 0; i <= compareWorkbookSheet.getLastRowNum(); i++) {
            Row row = compareWorkbookSheet.getRow(i);
            if (row.getLastCellNum() < 4) {
                throw new RuntimeException("请填写完整第" + (i + 1) + "行校验内容");
            }
            String sheetCode = row.getCell(0).getStringCellValue().trim();
            String indexLocation = row.getCell(1).getStringCellValue().trim(); //位置代码
            String indexCode = row.getCell(2).getStringCellValue().trim(); //指标代码
            String indexName = row.getCell(3).getStringCellValue().trim(); //指标名称
            List<String> indexLocationList;
            if (compareMap.containsKey(sheetCode)) {
                indexLocationList = compareMap.get(sheetCode);
            } else {
                indexLocationList = new ArrayList<>();
                compareMap.put(sheetCode, indexLocationList);
            }
            locationMap.put(indexLocation, indexCode);
            locationMap.put(indexLocation, indexName);
            indexLocationList.add(indexLocation);
        }
    }
}