package pers.lucas.lee.attachment2docx.utils;

import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddressList;

import java.util.ArrayList;
import java.util.Collection;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicInteger;

/**
 * Excel联动下拉选择工具类
 * 数据源存储在隐藏sheet中
 * 基于Apache POI 4.1.2
 */
public class ExcelDropdownUtil {
    /**
     * 序列号，用于EXCEL生成唯一名称，避免多次调用时名称重复问题
     */
    private static final AtomicInteger SERIAL = new AtomicInteger();

    /**
     * 创建下拉列表<br>
     * 数据源会自动存储在隐藏的sheet中
     *
     * @param workbook  Excel工作簿
     * @param sheetName 目标sheet名称
     * @param dataList  数据，选项列表
     * @param colIndex  下拉所在位置
     * @param startRow  开始行(0-based) 下拉范围行-起
     * @param endRow    结束行(0-based) 下拉范围行-止
     */
    public static void createDropDown(Workbook workbook, String sheetName, Collection<String> dataList,
                                      int colIndex, int startRow, int endRow) {
        if (dataList == null || dataList.isEmpty()) {
            return;
        }

        // 获取目标Sheet
        Sheet targetSheet = getTargetSheet(workbook, sheetName);

        int serialNum = SERIAL.getAndIncrement();

        // 创建隐藏sheet存储数据
        String hideSheetName = "_hide1_" + serialNum;
        Sheet hideSheet = createHideDataSheet(workbook, hideSheetName);

        DataValidationHelper validationHelper = targetSheet.getDataValidationHelper();

        // 写入选项到隐藏sheet的第一列
        setFirstLevelOptions(workbook, sheetName, colIndex, startRow, endRow, dataList, hideSheet, hideSheetName, validationHelper, targetSheet, serialNum);
    }


    /**
     * 创建两级联动下拉列表<br>
     * 数据源会自动存储在隐藏的sheet中
     *
     * @param workbook       Excel工作簿
     * @param sheetName      目标sheet名称
     * @param dataMap        联动数据Map，key为一级选项，value为对应的二级选项列表
     * @param firstColIndex  第一列索引(0-based) 一级下拉选项值所在列索引
     * @param secondColIndex 第二列索引(0-based) 二级下拉选项值所在列索引
     * @param startRow       开始行(0-based) 下拉范围行-起
     * @param endRow         结束行(0-based) 下拉范围行-止
     */
    public static void createTwoLevelCascade(Workbook workbook, String sheetName,
                                             Map<String, List<String>> dataMap,
                                             int firstColIndex, int secondColIndex,
                                             int startRow, int endRow) {
        if (dataMap == null || dataMap.isEmpty()) {
            return;
        }

        // 获取目标Sheet
        Sheet targetSheet = getTargetSheet(workbook, sheetName);

        int serialNum = SERIAL.getAndIncrement();

        // 创建隐藏sheet存储数据
        String hideSheetName = "_hide2_" + serialNum;
        Sheet hideSheet = createHideDataSheet(workbook, hideSheetName);

        DataValidationHelper validationHelper = targetSheet.getDataValidationHelper();


        // 写入一级选项到隐藏sheet的第一列
        List<String> firstLevelOptions = new ArrayList<>(dataMap.keySet());
        setFirstLevelOptions(workbook, sheetName, firstColIndex, startRow, endRow, firstLevelOptions, hideSheet, hideSheetName, validationHelper, targetSheet, serialNum);

        // 写入二级选项并创建命名区域
        int colIndex = 1;
        for (Map.Entry<String, List<String>> entry : dataMap.entrySet()) {
            String level1Key = entry.getKey();
            List<String> level2Values = entry.getValue();

            // 将二级选项写入隐藏sheet
            for (int i = 0; i < level2Values.size(); i++) {
                Row row = hideSheet.getRow(i);
                if (row == null) {
                    row = hideSheet.createRow(i);
                }
                row.createCell(colIndex).setCellValue(level2Values.get(i));
            }

            // 为每个一级选项的二级列表创建命名区域
            String safeName = createSafeName(level1Key);
            String nameStr = "SL" + serialNum + "_" + sanitizeSheetName(sheetName) + "_" + safeName;
            removeName(workbook, nameStr);

            Name name = workbook.createName();
            name.setNameName(nameStr);
            String columnLetter = getColumnLetter(colIndex);
            name.setRefersToFormula(hideSheetName + "!$" + columnLetter + "$1:$" +
                    columnLetter + "$" + Math.max(level2Values.size(), 1));

            colIndex++;
        }

        // 为每一行设置二级下拉列表
        String prefix = "SL" + serialNum + "_" + sanitizeSheetName(sheetName) + "_";
        for (int row = startRow; row <= endRow; row++) {
            String cellRef = getCellReference(firstColIndex, row);

            // 使用INDIRECT公式引用命名区域
            String formula = "INDIRECT(\"" + prefix + "\"&SUBSTITUTE(" + cellRef + ",\" \",\"_\"))";

            CellRangeAddressList secondAddressList = new CellRangeAddressList(
                    row, row, secondColIndex, secondColIndex);
            DataValidationConstraint secondConstraint = validationHelper.createFormulaListConstraint(formula);
            DataValidation secondValidation = validationHelper.createValidation(secondConstraint, secondAddressList);
            secondValidation.setShowErrorBox(true);
            secondValidation.setSuppressDropDownArrow(true);
            secondValidation.setEmptyCellAllowed(true);
            targetSheet.addValidationData(secondValidation);
        }

    }

    /**
     * 创建三级联动下拉列表<br>
     * 数据源会自动存储在隐藏的sheet中
     *
     * @param workbook       Excel工作簿
     * @param sheetName      目标sheet名称
     * @param dataMap        联动数据Map，key为一级选项，value为Map(key为二级选项，value为三级选项列表)
     * @param firstColIndex  第一列索引(0-based) 一级下拉选项值所在列索引
     * @param secondColIndex 第二列索引(0-based) 二级下拉选项值所在列索引
     * @param thirdColIndex  第三列索引(0-based) 三级下拉选项值所在列索引
     * @param startRow       开始行(0-based) 下拉范围行-起
     * @param endRow         结束行(0-based) 下拉范围行-止
     */
    public static void createThreeLevelCascade(Workbook workbook, String sheetName,
                                               Map<String, Map<String, List<String>>> dataMap,
                                               int firstColIndex, int secondColIndex, int thirdColIndex,
                                               int startRow, int endRow) {
        if (dataMap == null || dataMap.isEmpty() || dataMap.values().stream().allMatch(item -> item == null || item.isEmpty())) {
            return;
        }

        // 获取目标Sheet
        Sheet targetSheet = getTargetSheet(workbook, sheetName);

        int serialNum = SERIAL.getAndIncrement();

        // 创建隐藏sheet
        String hideSheetName = "_hide3_" + serialNum;
        Sheet hideSheet = createHideDataSheet(workbook, hideSheetName);

        DataValidationHelper validationHelper = targetSheet.getDataValidationHelper();

        // 处理一级选项
        List<String> firstLevelOptions = new ArrayList<>(dataMap.keySet());
        for (int i = 0; i < firstLevelOptions.size(); i++) {
            Row row = hideSheet.getRow(i);
            if (row == null) {
                row = hideSheet.createRow(i);
            }
            row.createCell(0).setCellValue(firstLevelOptions.get(i));
        }

        // 创建一级命名区域
        String firstNameStr = "FL" + serialNum + "_" + sanitizeSheetName(sheetName);
        removeName(workbook, firstNameStr);

        Name firstName = workbook.createName();
        firstName.setNameName(firstNameStr);
        firstName.setRefersToFormula(hideSheetName + "!$A$1:$A$" + Math.max(firstLevelOptions.size(), 1));

        // 设置一级下拉列表
        CellRangeAddressList firstAddressList = new CellRangeAddressList(
                startRow, endRow, firstColIndex, firstColIndex);
        DataValidationConstraint firstConstraint = validationHelper.createFormulaListConstraint(firstNameStr);
        DataValidation firstValidation = validationHelper.createValidation(firstConstraint, firstAddressList);
        firstValidation.setShowErrorBox(true);
        firstValidation.setSuppressDropDownArrow(true);
        targetSheet.addValidationData(firstValidation);

        // 处理二级和三级选项
        int colIndex = 1;
        String sheetPrefix = sanitizeSheetName(sheetName);

        for (Map.Entry<String, Map<String, List<String>>> level1Entry : dataMap.entrySet()) {
            String level1Key = level1Entry.getKey();
            Map<String, List<String>> level2Map = level1Entry.getValue();

            List<String> level2Options = new ArrayList<>(level2Map.keySet());

            // 写入二级选项
            for (int i = 0; i < level2Options.size(); i++) {
                Row row = hideSheet.getRow(i);
                if (row == null) {
                    row = hideSheet.createRow(i);
                }
                row.createCell(colIndex).setCellValue(level2Options.get(i));
            }

            // 创建二级命名区域
            String level2SafeName = createSafeName(level1Key);
            String level2NameStr = "SL" + serialNum + "_" + sheetPrefix + "_" + level2SafeName;
            removeName(workbook, level2NameStr);

            Name level2Name = workbook.createName();
            level2Name.setNameName(level2NameStr);
            String columnLetter = getColumnLetter(colIndex);
            level2Name.setRefersToFormula(hideSheetName + "!$" + columnLetter + "$1:$" +
                    columnLetter + "$" + Math.max(level2Options.size(), 1));
            colIndex++;

            // 处理三级选项
            for (Map.Entry<String, List<String>> level2Entry : level2Map.entrySet()) {
                String level2Key = level2Entry.getKey();
                List<String> level3Options = level2Entry.getValue();

                // 写入三级选项
                for (int i = 0; i < level3Options.size(); i++) {
                    Row row = hideSheet.getRow(i);
                    if (row == null) {
                        row = hideSheet.createRow(i);
                    }
                    row.createCell(colIndex).setCellValue(level3Options.get(i));
                }

                // 创建三级命名区域
                String level3SafeName = createSafeName(level1Key + "_" + level2Key);
                String level3NameStr = "TL" + serialNum + "_" + sheetPrefix + "_" + level3SafeName;
                removeName(workbook, level3NameStr);

                Name level3Name = workbook.createName();
                level3Name.setNameName(level3NameStr);
                columnLetter = getColumnLetter(colIndex);
                level3Name.setRefersToFormula(hideSheetName + "!$" + columnLetter + "$1:$" +
                        columnLetter + "$" + Math.max(level3Options.size(), 1));
                colIndex++;
            }
        }

        // 设置二级下拉列表
        String prefix2 = "SL" + serialNum + "_" + sheetPrefix + "_";
        for (int row = startRow; row <= endRow; row++) {
            String cellRef = getCellReference(firstColIndex, row);
            String formula = "INDIRECT(\"" + prefix2 + "\"&SUBSTITUTE(" + cellRef + ",\" \",\"_\"))";

            CellRangeAddressList secondAddressList = new CellRangeAddressList(
                    row, row, secondColIndex, secondColIndex);
            DataValidationConstraint secondConstraint = validationHelper.createFormulaListConstraint(formula);
            DataValidation secondValidation = validationHelper.createValidation(secondConstraint, secondAddressList);
            secondValidation.setShowErrorBox(true);
            secondValidation.setSuppressDropDownArrow(true);
            secondValidation.setEmptyCellAllowed(true);
            targetSheet.addValidationData(secondValidation);
        }

        // 设置三级下拉列表
        String prefix3 = "TL" + serialNum + "_" + sheetPrefix + "_";
        for (int row = startRow; row <= endRow; row++) {
            String cell1Ref = getCellReference(firstColIndex, row);
            String cell2Ref = getCellReference(secondColIndex, row);
            String formula = "INDIRECT(\"" + prefix3 + "\"&SUBSTITUTE(" + cell1Ref +
                    ",\" \",\"_\")&\"_\"&SUBSTITUTE(" + cell2Ref + ",\" \",\"_\"))";

            CellRangeAddressList thirdAddressList = new CellRangeAddressList(
                    row, row, thirdColIndex, thirdColIndex);
            DataValidationConstraint thirdConstraint = validationHelper.createFormulaListConstraint(formula);
            DataValidation thirdValidation = validationHelper.createValidation(thirdConstraint, thirdAddressList);
            thirdValidation.setShowErrorBox(true);
            thirdValidation.setSuppressDropDownArrow(true);
            thirdValidation.setEmptyCellAllowed(true);
            targetSheet.addValidationData(thirdValidation);
        }

    }


    /**
     * 获取指定名称的sheet，未获取到则创建
     *
     * @param workbook  Excel工作簿
     * @param sheetName 目标sheet名称
     * @return 目标sheet
     */
    private static Sheet getTargetSheet(Workbook workbook, String sheetName) {
        Sheet targetSheet = workbook.getSheet(sheetName);
        if (targetSheet == null) {
            targetSheet = workbook.createSheet(sheetName);
        }
        return targetSheet;
    }

    /**
     * 创建存储数据的隐藏sheet
     *
     * @param workbook      Excel工作簿
     * @param hideSheetName 隐藏sheet名称
     * @return 隐藏sheet
     */
    private static Sheet createHideDataSheet(Workbook workbook, String hideSheetName) {
        Sheet hideSheet = workbook.getSheet(hideSheetName);
        if (hideSheet == null) {
            hideSheet = workbook.createSheet(hideSheetName);
            // 设置sheet为隐藏
            workbook.setSheetHidden(workbook.getSheetIndex(hideSheet), true);
        }
        return hideSheet;
    }

    /**
     * 创建一级下拉
     *
     * @param workbook          Excel工作簿
     * @param sheetName         目标sheet名称
     * @param firstColIndex     第一列索引(0-based) 一级下拉选项值所在列索引
     * @param startRow          开始行(0-based) 下拉范围行-起
     * @param endRow            结束行(0-based) 下拉范围行-止
     * @param firstLevelOptions 数据，选项列表
     * @param hideSheet         存储数据的隐藏sheet
     * @param hideSheetName     存储数据的隐藏sheet名称
     * @param validationHelper  处理Excel数据验证的对象
     * @param targetSheet       目标sheet
     */
    private static void setFirstLevelOptions(Workbook workbook, String sheetName, int firstColIndex, int startRow, int endRow, Collection<String> firstLevelOptions, Sheet hideSheet, String hideSheetName, DataValidationHelper validationHelper, Sheet targetSheet, int serialNum) {
        int i = 0;
        for (String value : firstLevelOptions) {
            Row row = hideSheet.getRow(i);
            if (row == null) {
                row = hideSheet.createRow(i);
            }
            row.createCell(0).setCellValue(value);
            i++;
        }

        // 创建一级下拉列表的命名区域
        String firstNameStr = "FL" + serialNum + "_" + sanitizeSheetName(sheetName);
        removeName(workbook, firstNameStr);

        Name firstName = workbook.createName();
        firstName.setNameName(firstNameStr);
        firstName.setRefersToFormula(hideSheetName + "!$A$1:$A$" + Math.max(firstLevelOptions.size(), 1));

        // 在目标sheet设置一级下拉列表
        CellRangeAddressList firstAddressList = new CellRangeAddressList(
                startRow, endRow, firstColIndex, firstColIndex);
        DataValidationConstraint firstConstraint = validationHelper.createFormulaListConstraint(firstNameStr);
        DataValidation firstValidation = validationHelper.createValidation(firstConstraint, firstAddressList);
        firstValidation.setShowErrorBox(true);
        firstValidation.setSuppressDropDownArrow(true);
        targetSheet.addValidationData(firstValidation);
    }


    /**
     * 清理sheet名称用于命名
     */
    private static String sanitizeSheetName(String name) {
        return name.replaceAll("[^a-zA-Z0-9]", "");
    }

    /**
     * 移除已存在的命名区域
     */
    private static void removeName(Workbook workbook, String nameName) {
        Name name = workbook.getName(nameName);
        if (name != null) {
            workbook.removeName(name);
        }
    }

    /**
     * 创建安全的命名区域名称
     */
    private static String createSafeName(String original) {
        if (original == null || original.isEmpty()) {
            return "Empty";
        }

        String safeName = original.replace(" ", "_")
                .replaceAll("[^a-zA-Z0-9_\u4e00-\u9fa5]", "_");

        if (safeName.matches("^[0-9].*")) {
            safeName = "N" + safeName;
        }

        if (safeName.length() > 150) {
            safeName = safeName.substring(0, 150);
        }

        return safeName;
    }

    /**
     * 获取列字母表示
     */
    private static String getColumnLetter(int columnIndex) {
        StringBuilder columnName = new StringBuilder();
        while (columnIndex >= 0) {
            columnName.insert(0, (char) ('A' + (columnIndex % 26)));
            columnIndex = columnIndex / 26 - 1;
        }
        return columnName.toString();
    }

    /**
     * 获取单元格引用
     */
    private static String getCellReference(int colIndex, int rowIndex) {
        return getColumnLetter(colIndex) + (rowIndex + 1);
    }
}