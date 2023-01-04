package com.example.demo;

import javafx.fxml.FXML;
import javafx.fxml.FXMLLoader;
import javafx.scene.Scene;
import javafx.scene.control.Alert;
import javafx.stage.DirectoryChooser;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import javafx.stage.Window;
import org.apache.commons.collections4.MultiValuedMap;
import org.apache.commons.collections4.multimap.ArrayListValuedHashMap;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.text.NumberFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class HelloController {
    static final String compareFile = "C:\\compare.xls";
    static final String[] titles = {"报表代码", "指标位置", "指标代码", "指标名称", "本期数据值", "上期数据值", "比上期（万元）", "比上期（%）", "备注"};
    Map<String, List<String>> compareMap = new HashMap<>();
    Map<String, List<Double>> previousMap = new HashMap<>();
    Map<String, List<Double>> currentMap = new HashMap<>();
    MultiValuedMap<String, String> locationMap = new ArrayListValuedHashMap<>();

    @FXML
    protected void onUploadButtonClick(){
        FileChooser chooser = new FileChooser();
        //设置文件上传类型
        FileChooser.ExtensionFilter extFilter = new FileChooser.ExtensionFilter("excel files (*.xls)", "*.xls","*.xlsx");
        chooser.getExtensionFilters().add(extFilter);
        chooser.setTitle("上传校验文件");
        File file = chooser.showOpenDialog(Window.impl_getWindows().next());
        try {
            doHandlerCompareFile(file);
        } catch (Exception e) {
            Alert alert = new Alert(Alert.AlertType.WARNING);
            alert.titleProperty().set("警告");
            alert.headerTextProperty().set(e.getMessage());
            alert.showAndWait();
        }
    }

    @FXML
    protected void onDownloadButtonClick(){

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

            File resultFile = new File("C:\\数据比对结果.xls");
            wb.write(resultFile);

            Alert alert = new Alert(Alert.AlertType.INFORMATION);
            alert.titleProperty().set("解析完成");
            alert.headerTextProperty().set("已生成结果文件 C:\\数据比对结果.xls");
            alert.showAndWait();
        } catch (Exception e) {
            Alert alert = new Alert(Alert.AlertType.WARNING);
            alert.titleProperty().set("警告");
            alert.headerTextProperty().set(e.getMessage());
            alert.showAndWait();
        }
    }

    @FXML
    protected void onCompareButtonClick() throws IOException {
        // 创建新的stage
        Stage secondStage = new Stage();
        FXMLLoader fxmlLoader = new FXMLLoader(HelloApplication.class.getResource("second-view.fxml"));
        Scene secondScene = new Scene(fxmlLoader.load(), 320, 240);
        secondStage.setTitle("校验文件");
        secondStage.setScene(secondScene);
        secondStage.show();
    }

    @FXML
    protected void onPreviousButtonClick() {
        String property = System.getProperty("user.dir");
        System.out.println(property);
        DirectoryChooser chooser = new DirectoryChooser();
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
        doHandlerCompareFile(null);

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


    }

    private void doHandlerCompareFile(File file) throws Exception {
        if (file == null) {
            file = new File(compareFile);
        }else {
            FileReader fileReader = new FileReader(file);
            FileWriter fileWriter = new FileWriter(compareFile);
            fileWriter.write(fileReader.read());
        }
        if (file == null || !file.exists()) {
            throw new RuntimeException("请手动上传校验文件或将校验文件放在指定位置 C:\\compare.xls");
        }

        // 将上次内容清空
        compareMap.clear();

        Workbook compareWorkbook = new HSSFWorkbook(new FileInputStream(file));
        Sheet compareWorkbookSheet = compareWorkbook.getSheetAt(0);
        for (int i = 0; i <= compareWorkbookSheet.getLastRowNum(); i++) {
            Row row = compareWorkbookSheet.getRow(i);
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