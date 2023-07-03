package com.example.demo;

import com.example.util.CopySheetUtils;
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
import org.apache.commons.io.IOUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.NumberFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class HelloController {
    //private static final Logger logger = LoggerFactory.getLogger(TestController.class);
    //private static final String COMPARE_FILE = "校验文件.xls";
    //private static final String COMPARE_TEMPLATE_FILE = "compare_template.xls";
    //static final String[] titles = {"报表代码", "指标位置", "指标代码", "指标名称", "本期数据值", "上期数据值", "比上期（万元）", "比上期（%）", "备注"};
    static final String BASE_PATH = System.getProperty("user.dir") + File.separator;
    private static final String RESULT_FILE = "解析结果.zip";
    private static final String RESULT_FILE_SUFFIX = "_比对结果.xls";
    private static final Map<String, File> fileMap = new HashMap<>(2);

    /*@FXML
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
    }*/

    void alert(Alert.AlertType alertType, String title, String text) {
        Alert alert = new Alert(alertType);
        alert.titleProperty().set(title);
        alert.headerTextProperty().set(text);
        alert.showAndWait();
    }

    @FXML
    protected void onStartButtonClick() {
        try {
            if (fileMap.get("previousFile") == null) {
                throw new RuntimeException("请上传上期文件夹");
            }
            if (fileMap.get("currentFile") == null) {
                throw new RuntimeException("请上传当期文件夹");
            }
            FileOutputStream outputStream = null;
            try {
                for (File currentMultipartFile : fileMap.get("currentFile").listFiles()) {
                    // 根据当月文件获取报表代码
                    String filename = FilenameUtils.getBaseName(currentMultipartFile.getName());
                    String code = filename.substring(0, filename.indexOf("-"));
                    // 20230607新增 人行报表以机构号开头，报表名结尾
                    if (filename.matches("^(\\d{10})#.+")) {
                        String[] split = filename.split("#");
                        code = split[0] + split[split.length - 1];
                    }
                    // 根据当月报表代码查找上月文件
                    for (File previousMultipartFile : fileMap.get("previousFile").listFiles()) {
                        if (matchCode(FilenameUtils.getBaseName(previousMultipartFile.getName()), code)) {
                            HSSFWorkbook currentWorkbook = null;
                            HSSFWorkbook previousWorkbook = null;
                            // 开始比对
                            try {
                                currentWorkbook = new HSSFWorkbook(new FileInputStream(currentMultipartFile));
                                previousWorkbook = new HSSFWorkbook(new FileInputStream(previousMultipartFile));
                                // 原当期sheet
                                HSSFSheet currentOriginalSheet = currentWorkbook.getSheetAt(0);
                                // 原上期sheet
                                HSSFSheet previousOriginalSheet = previousWorkbook.getSheetAt(0);
                                HSSFSheet previousCloneSheet = currentWorkbook.createSheet("上期数据");
                                CopySheetUtils.copySheets(previousCloneSheet, previousOriginalSheet);
                                // 比上期差值sheet
                                HSSFSheet diffSheet = currentWorkbook.cloneSheet(0);
                                // 比上期百分比sheet
                                HSSFSheet ratioSheet = currentWorkbook.cloneSheet(0);

                                currentWorkbook.setSheetName(0, "当期数据");
                                currentWorkbook.setSheetName(2, "比上期差值");
                                currentWorkbook.setSheetName(3, "比上期百分比");
                                currentWorkbook.setActiveSheet(2);

                                // 1.找到当月文件第一个数字单元格的位置
                            /*CellAddress address = getCellAddress(currentOriginalSheet);
                            if (address == null) {
                                continue;
                            }*/
                                // 2.基于当月文件，从第一个数字单元格的位置开始进行比对，比对时直接在当月文件上修改
                                for (int i = 0; i < diffSheet.getLastRowNum(); i++) {
                                    HSSFRow currentRow = diffSheet.getRow(i);
                                    HSSFRow previousRow = previousOriginalSheet.getRow(i);
                                    if (currentRow == null || previousRow == null) {
                                        continue;
                                    }
                                    for (int j = 0; j < currentRow.getLastCellNum(); j++) {
                                        HSSFCell cell = currentRow.getCell(j);
                                        Double currentValue = getCellValue(cell);
                                        HSSFCell previousCell = previousRow.getCell(j);
                                        Double previousValue = getCellValue(previousCell);
                                        if (currentValue == null || previousValue == null) {
                                            continue;
                                        }
                                        // 比对之后赋值
                                        if (cell.getCellTypeEnum() == CellType.FORMULA) {
                                            cell.setCellFormula(null);
                                            cell.setCellType(CellType.STRING);
                                        }
                                        cell.setCellValue(currentValue - previousValue);
                                        HSSFCell cell2 = ratioSheet.getRow(i).getCell(j);
                                        if (previousValue != 0) {
                                            NumberFormat numberFormat = NumberFormat.getInstance();
                                            numberFormat.setMinimumFractionDigits(2);
                                            cell2.setCellType(CellType.STRING);
                                            cell2.setCellValue(numberFormat.format((currentValue - previousValue) / Math.abs(previousValue) * 100) + "%");
                                        }
                                        // 修改单元格样式
                                        CellStyle cellStyle = getCellStyle(currentWorkbook, cell, currentValue, previousValue);
                                        cell.setCellStyle(cellStyle);
                                        cell2.setCellStyle(cellStyle);
                                    }
                                }
                                currentWorkbook.write(new File(filename + "_" + RESULT_FILE_SUFFIX));
                            } catch (Exception e) {
                                //logger.error("", e);
                            } finally {
                                IOUtils.closeQuietly(currentWorkbook);
                                IOUtils.closeQuietly(previousWorkbook);
                                break;
                            }
                        }
                    }
                }

                Alert alert = new Alert(Alert.AlertType.INFORMATION);
                alert.titleProperty().set("解析完成");
                //alert.headerTextProperty().set("已生成结果文件 " + resultFileName);
                alert.showAndWait();
            } catch (Exception e) {
                alert(Alert.AlertType.WARNING, "警告", e.getMessage());
            }
        }
    }
    /*@FXML
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
    }*/

    @FXML
    protected void onPreviousButtonClick() {
        DirectoryChooser chooser = new DirectoryChooser();
        chooser.setTitle("上传上期文件夹");
        File file = chooser.showDialog(Window.impl_getWindows().next());
        //File file = chooser.showDialog(Window.getWindows().get(0));
        if (file != null) {
            fileMap.put("previousFile", file);
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
            fileMap.put("currentFile", file);
        }
    }

    private boolean matchCode(String baseName, String code) {
        if (baseName.startsWith(code)) {
            return true;
        }
        String[] split = baseName.split("#");
        if (baseName.matches("^(\\d{10})#.+") && StringUtils.equals(code, split[0] + split[split.length - 1])) {
            return true;
        }
        return false;
    }

    /**
     * 当期和前期进行比较
     * 比上期变动幅度大于100% 棕黄色
     * 比上期变动幅度等于100% 黄色
     * 比上期变动幅度小于-100% 红色
     * 本期数据为0，上期不等于0 绿色
     * 本期数据不等于0，上期数据为0 蓝色
     */
    private CellStyle getCellStyle(HSSFWorkbook currentWorkbook, HSSFCell cell, Double currentValue, Double previousValue) {
        //String remarks;
        CellStyle cellStyle = currentWorkbook.createCellStyle();
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cellStyle.setAlignment(cell.getCellStyle().getAlignmentEnum());
        cellStyle.setBorderBottom(cell.getCellStyle().getBorderBottomEnum());
        cellStyle.setBorderLeft(cell.getCellStyle().getBorderLeftEnum());
        cellStyle.setBorderRight(cell.getCellStyle().getBorderRightEnum());
        cellStyle.setBorderTop(cell.getCellStyle().getBorderTopEnum());
        cellStyle.setBottomBorderColor(cell.getCellStyle().getBottomBorderColor());
        cellStyle.setLeftBorderColor(cell.getCellStyle().getLeftBorderColor());
        cellStyle.setRightBorderColor(cell.getCellStyle().getRightBorderColor());
        cellStyle.setTopBorderColor(cell.getCellStyle().getTopBorderColor());

        if (previousValue != 0 && currentValue == 0) {
            cellStyle.setFillForegroundColor(HSSFColor.HSSFColorPredefined.GREEN.getIndex());
            //remarks = "本期数据为0，上期不等于0";
            //logger.info("{}设置绿色", index);
        } else if (previousValue == 0 && currentValue != 0) {
            cellStyle.setFillForegroundColor(HSSFColor.HSSFColorPredefined.LIGHT_BLUE.getIndex());
            //remarks = "本期数据不等于0，上期数据为0";
            //logger.info("{}设置蓝色", index);
        }

        if (previousValue == 0 || currentValue == 0) {
            return cellStyle;
        }

        double ratio = (currentValue - previousValue) / Math.abs(previousValue);
        if (ratio > 1) {
            cellStyle.setFillForegroundColor(HSSFColor.HSSFColorPredefined.ORANGE.getIndex());
            //remarks = "比上期变动幅度大于100%";
            //logger.info("{}设置棕黄色", index);
        } else if (ratio == 1) {
            cellStyle.setFillForegroundColor(HSSFColor.HSSFColorPredefined.YELLOW.getIndex());
            //remarks = "比上期变动幅度等于100%";
            //logger.info("{}设置黄色", index);
        } else if (ratio < -1) {
            cellStyle.setFillForegroundColor(HSSFColor.HSSFColorPredefined.RED.getIndex());
            //remarks = "比上期变动幅度小于-100%";
            //logger.info("{}设置红色", index);
        }
        return cellStyle;
    }

    private Double getCellValue(HSSFCell cell) {
        if (cell == null) {
            return null;
        }
        Double currentValue = null;
        if (CellType.NUMERIC == cell.getCellTypeEnum()) {
            currentValue = cell.getNumericCellValue();
        } else if (CellType.STRING == cell.getCellTypeEnum()) {
            try {
                currentValue = Double.valueOf(cell.getStringCellValue());
            } catch (Exception e) {
                //logger.error(cell.getStringCellValue() + "转换为数字格式失败");
            }
        } else if (CellType.FORMULA == cell.getCellTypeEnum()) {
            currentValue = cell.getNumericCellValue();
        }
        return currentValue;
    }

    /*private void doHandlerDirectory(File file, Map<String, List<Double>> valueMap) throws Exception {
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
    }*/
}