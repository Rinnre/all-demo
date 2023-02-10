package com.wj.excel.util;

import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.stream.Collectors;

/**
 * Created by IntelliJ IDEA
 *
 * @author wj
 * @date 2022/10/28 14:39
 */
public class ExcelUtil3 {

    private static DateTimeFormatter dateTimeFormatter;
    private static DateTimeFormatter dateFormatter;
    private static DateTimeFormatter timeFormatter;

    static {
        dateTimeFormatter = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss", Locale.CHINA);
        dateFormatter = DateTimeFormatter.ofPattern("yyyy-MM-dd", Locale.CHINA);
        timeFormatter = DateTimeFormatter.ofPattern("HH:mm:ss", Locale.CHINA);
    }

    /**
     * 反射来填充excelEntity中的数据
     * @param excelEntity
     * @param constantMap
     * @param objects
     */
    public static void fillExceleEntity(ExcelEntity excelEntity,Map<String,Map<String, String>> constantMap,List<?> objects){
        Map<String, List<Object>> tableDataList = new HashMap<>();
        LinkedHashMap<String, String> tableFiled = excelEntity.getTableFiled();
        Set<String> filedSet = tableFiled.keySet();
        for (String filed : filedSet) {
            List<Object> data = new ArrayList<>();
            objects.forEach(object -> {
                String methodName = "get"+filed.replaceFirst(filed.substring(0,1),filed.substring(0,1).toUpperCase());
                Object value = null;
                try {
                    Method method = object.getClass().getMethod(methodName);
                    value = method.invoke(object);
                    try {
                        value = Boolean.parseBoolean(value.toString()) ;
                        value = constantMap.get(filed).get(value.toString());
                    } catch (Exception e) {

                    }
                } catch (NoSuchMethodException e) {
//                    logActionManager.insertEntityFgn(ErrorMark.EXCEPTION,Module.DeviceSvrQue,messageId,CurrentLineInfo.getDebugInfo(e));
                } catch (IllegalAccessException e) {
//                    logActionManager.insertEntityFgn(ErrorMark.EXCEPTION,Module.DeviceSvrQue,messageId,CurrentLineInfo.getDebugInfo(e));
                } catch (InvocationTargetException e) {
//                    logActionManager.insertEntityFgn(ErrorMark.EXCEPTION,Module.DeviceSvrQue,messageId,CurrentLineInfo.getDebugInfo(e));
                }
                data.add(value);
            });
            tableDataList.put(filed,data);
        }
        excelEntity.setTableDataList(tableDataList);
    }

    /**
     * 多个sheet
     *
     * @param workbook
     * @param excelEntities
     */
    public static void getExcelTable(XSSFWorkbook workbook, List<ExcelEntity> excelEntities) {
        excelEntities.forEach(excelEntity ->
                getExcelTable(workbook, excelEntity)
        );
    }

    /**
     * 单个sheet
     *
     * @param excelEntity
     * @return
     */
    public static void getExcelTable(XSSFWorkbook workbook, ExcelEntity excelEntity) {

        // 获取表名、属性名、数据
        Map<String, List<Object>> tableDataList = excelEntity.getTableDataList();
        String tableName = excelEntity.getTableName();
        LinkedHashMap<String, String> tableFiled = excelEntity.getTableFiled();
        boolean isStatistics = excelEntity.getStatistics();
        int realRow = 0;
        List<Integer> colList = new ArrayList<>();
        Map<String, Object> filedUnit = excelEntity.getFiledUnit();

        if (tableDataList == null || tableDataList.size() <= 0) {
            throw new RuntimeException(tableName + ":excel data is empty!");
        }

        // 获取sheet
        XSSFSheet sheet = workbook.createSheet(tableName);

        //设置表的默认列宽
        sheet.setDefaultColumnWidth(20);

        // 创建表头以及属性行
        XSSFRow headRow = sheet.createRow(realRow++);
        XSSFCell headCell = headRow.createCell(0);
        headCell.setCellValue(tableName);

        CellStyle headStyle = headStyle(workbook);
        headCell.setCellStyle(headStyle);
        // 合并列数
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, tableFiled.size() - 1));

        headRow = sheet.createRow(realRow++);
        Set<String> filedSet = tableFiled.keySet();
        int headCol = 0;
        for (String filed : filedSet) {
            XSSFCell cell = headRow.createCell(headCol);
            if ((!filedSet.iterator().next().equals(filed)) && null != filedUnit) {
                cell.setCellValue(tableFiled.get(filed) + "/" + filedUnit.get("y").toString().split("/")[1]);
            } else {
                cell.setCellValue(tableFiled.get(filed));
            }
            headCol++;
            cell.setCellStyle(headStyle);
        }


        //创建内容行样式
        CellStyle contentStyle = contentStyle(workbook);

        // 添加数据
        // 获取总行数
        int rowNum = tableDataList.get(tableDataList.keySet().iterator().next()).size();
        int colNum = tableDataList.size();
        for (int row = realRow; row < rowNum + realRow; row++) {
            XSSFRow dataRow = sheet.createRow(row);
            int col = 0;
            for (String filed : tableFiled.keySet()) {
                XSSFCell cell = dataRow.createCell(col);
                // 获取数据
                Object objectData = tableDataList.get(filed).get(row - realRow);
                if (objectData != null) {
                    try {
                        cell.setCellValue(Double.parseDouble(objectData.toString()));
                        colList.add(col);
                    } catch (NumberFormatException e) {
                        cell.setCellValue(objectData.toString());
                    }
                } else {
                    cell.setCellValue("-");
                }
                col++;
                cell.setCellStyle(contentStyle);
            }
        }

        // 数据列去重
        colList = colList.stream().distinct().collect(Collectors.toList());

        // 统计数据列总数据
        if (excelEntity.getStatistics()) {
            int lastRowNum = sheet.getLastRowNum();
            XSSFRow totalRow = sheet.createRow(lastRowNum + 1);
            XSSFCell totalRowCell = totalRow.createCell(0);
            totalRowCell.setCellValue("total");
            totalRowCell.setCellStyle(contentStyle);
            for (int col = 0; col < colList.size(); col++) {
                int startCol = 65;
                XSSFCell cell = totalRow.createCell(colList.get(col));
                startCol += colList.get(col);
                String formula = "sum(" + ((char) startCol) + (realRow + 1) + ":" + ((char) startCol) + (lastRowNum + 1) + ")";
                cell.setCellFormula(formula);
                cell.setCellStyle(contentStyle);
            }
        }

        // 生成统计图
        if (excelEntity.getExcelChartTypeEnum() != null) {
            colList = colList.stream().distinct().collect(Collectors.toList());
            switch (excelEntity.getExcelChartTypeEnum()) {
                case LINE_CHAR:
                    tableFiled.remove(tableFiled.keySet().iterator().next());
                    ExcelImgUtil.createSheetLineChart(workbook.getSheet(excelEntity.getTableName()), excelEntity.getTableName(), tableFiled.values().stream().collect(Collectors.toList()), null != filedUnit ? filedUnit.get("x").toString() : "", null != filedUnit ? filedUnit.get("y").toString() : "", 2, tableDataList.get(filedSet.iterator().next()).size(), 0, 0, colList);
                    break;
                case PIE:
                    ExcelImgUtil.createSheetPieChart(workbook.getSheet(excelEntity.getTableName()), excelEntity.getTableName(), 2, tableDataList.get(filedSet.iterator().next()).size() + 1, 0, 0, colList, isStatistics);
                    break;
                case BAR_CHAR:
                    ExcelImgUtil.createSheetBarChart(workbook.getSheet(excelEntity.getTableName()), excelEntity.getTableName(), null != filedUnit ? filedUnit.get("x").toString() : "", null != filedUnit ? filedUnit.get("y").toString() : "", 2, tableDataList.get(filedSet.iterator().next()).size() + 1, 0, 0, colList, isStatistics);
                    break;
                default:
                    break;
            }
        }

    }


    /**
     * 创建内容行样式
     *
     * @param wb 工作簿
     * @return
     */
    public static CellStyle contentStyle(Workbook wb) {
        CellStyle contentStyle = wb.createCellStyle();
        Font contentFont = wb.createFont();
        contentFont.setFontName("宋体");
        contentFont.setColor(HSSFFont.COLOR_NORMAL);
        contentFont.setFontHeightInPoints((short) 11);

        contentStyle.setAlignment(HorizontalAlignment.CENTER);
        contentStyle.setFont(contentFont);
        return contentStyle;
    }


    /**
     * 创建标题行样式
     *
     * @param wb
     * @return
     */
    public static CellStyle headStyle(Workbook wb) {
        // 创建样式对象
        CellStyle headStyle = wb.createCellStyle();
        // 创建字体
        Font headFont = wb.createFont();
        headFont.setFontName("微软雅黑");
        headFont.setBold(true);
        headFont.setColor(HSSFFont.COLOR_NORMAL);
        headFont.setFontHeightInPoints((short) 11);
        headStyle.setAlignment(HorizontalAlignment.CENTER);
        headStyle.setFont(headFont);
        return headStyle;
    }

}
