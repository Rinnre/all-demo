package com.wj.poi.utils;

import com.alibaba.fastjson.JSONObject;
import com.alibaba.fastjson.TypeReference;

import com.wj.poi.entity.ExcelEntity;

import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFFooter;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;

import org.apache.poi.xssf.usermodel.*;


import java.io.*;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
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
     * 多个sheet
     *
     * @param workbook
     * @param excelEntities
     */
    public static void getExcelTable(XSSFWorkbook workbook, List<ExcelEntity> excelEntities) {
        excelEntities.forEach(excelEntity -> {
                    try {
                        getExcelTable(workbook, excelEntity);
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                }
        );
    }


    /**
     * 得到excel表
     *
     * @param workbook    工作簿
     * @param excelEntity excel实体
     */
    public static void getExcelTable(XSSFWorkbook workbook, ExcelEntity excelEntity) throws IOException {

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

        // 设置页眉
        Header header = sheet.getHeader();
        String logo = "KEEYODA" + "\n" + "企物能源管理系统";
        header.setCenter(logo);
        Footer footer = sheet.getFooter();
        footer.setCenter(logo);
        //设置表的默认列宽
        sheet.setDefaultColumnWidth(15);


        // 创建表头以及属性行
        XSSFRow headRow = sheet.createRow(realRow++);
        XSSFCell headCell = headRow.createCell(0);
        headCell.setCellValue(tableName);
        CellStyle headStyle = headStyle(workbook, false, false);
        headCell.setCellStyle(headStyle);
        // 表头列数合并
        CellRangeAddress tableHeadCell = new CellRangeAddress(realRow - 1, realRow - 1,
                0, tableFiled.size() - 1 > 4 ? tableFiled.size() - 1 : 4);

        sheet.addMergedRegion(tableHeadCell);
        if (null != filedUnit) {
            int cellCol = 0;
            String[] unitSplit =filedUnit.get("y").toString().split("/");
            String unit ="";
            if(unitSplit.length<=2){
                unit= filedUnit.get("y").toString().split("/")[1];
            }else{
                unit= filedUnit.get("y").toString().split("/")[1]+"/"+
                        filedUnit.get("y").toString().split("/")[2];
            }
            XSSFCell unitCell, timeCell = null;
            // 增加单位、时间行
            XSSFRow otherInfoRow = sheet.createRow(realRow++);
            XSSFCell nullCell = otherInfoRow.createCell(cellCol++);
            nullCell.setCellValue("");
            nullCell.setCellStyle(headStyle);
            timeCell = otherInfoRow.createCell(cellCol++);
            timeCell.setCellValue("时间");
            timeCell.setCellStyle(headStyle);
            timeCell = otherInfoRow.createCell(cellCol++);
            unitCell = otherInfoRow.createCell(cellCol++);
            unitCell.setCellValue("单位");
            unitCell.setCellStyle(headStyle);
            unitCell = otherInfoRow.createCell(cellCol++);


            unitCell.setCellValue(unit);
            unitCell.setCellStyle(headStyle(workbook, true, false));
            if (null != excelEntity.getTableHeadTime() && null != timeCell) {
                timeCell.setCellValue(excelEntity.getTableHeadTime());
                timeCell.setCellStyle(headStyle(workbook, true, false));
            }

            // 补充数据列
            for (; cellCol < tableFiled.size(); cellCol++) {
                XSSFCell cell = otherInfoRow.createCell(cellCol);
                cell.setCellValue("");
                cell.setCellStyle(headStyle);
            }
        }


//        // 插入logo
//        ClassPathResource classPathResource = new ClassPathResource("static/logo.png");
//        InputStream inputStream = classPathResource.getInputStream();
//        byte[] bytes = IOUtils.toByteArray(inputStream);
//        @SuppressWarnings("static-access")
//        int pictureIdx = workbook.addPicture(bytes, workbook.PICTURE_TYPE_PNG);
//        XSSFCreationHelper creationHelper = workbook.getCreationHelper();
//        XSSFDrawing drawingPatriarch = sheet.createDrawingPatriarch();
//        XSSFClientAnchor clientAnchor = creationHelper.createClientAnchor();
//        // 图片坐标
//        clientAnchor.setCol1(0);
//        clientAnchor.setCol2(2);
//        clientAnchor.setRow1(0);
//        clientAnchor.setRow2(realRow - 1);
//        XSSFPicture picture = drawingPatriarch.createPicture(clientAnchor, pictureIdx);
//        picture.resize(1, 1);

        headRow = sheet.createRow(realRow++);

        // 冻结表头
        sheet.createFreezePane(1, realRow, 1, realRow);
        Set<String> filedSet = tableFiled.keySet();
        int headCol = 0;
        for (String filed : filedSet) {
            XSSFCell cell = headRow.createCell(headCol);
            cell.setCellValue(tableFiled.get(filed));
            headCol++;
            cell.setCellStyle(headStyle(workbook, true, false));
        }
        // 补充列数
        for (; headCol < 5; headCol++) {
            XSSFCell cell = headRow.createCell(headCol);
            cell.setCellValue("");
            cell.setCellStyle(headStyle(workbook, true, false));
        }


        // 添加数据  统计第一行以及最后一行不为空的数据位置
        List<String> formulaNumList = new ArrayList<>();
        for (String filed : tableFiled.keySet()) {
            if (!"time".equals(filed)) {
                int startNum = 0, endNum = 0;
                List<Object> objects = tableDataList.get(filed);
                for (int i = 0; i < objects.size(); i++) {
                    if (null != objects.get(i)) {
                        startNum = i;
                        break;
                    }
                }
                for (int i = objects.size() - 1; i >= 0; i--) {
                    if (null != objects.get(i)) {
                        endNum = i;
                        break;
                    }
                }
                formulaNumList.add(endNum + "-" + startNum);
            }
        }

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
                        if ("time".equals(filed)) {
                            cell.setCellValue(Double.parseDouble(objectData.toString()));
                        }else{
                            cell.setCellValue(Double.parseDouble(objectData.toString()));
                            colList.add(col);
                        }
                    } catch (NumberFormatException e) {
                        if ("time".equals(filed)) {
                            objectData = dateToStringIfIsDate(objectData);
                        }
                        cell.setCellValue(objectData.toString());
                    }
                } else {
                    String s = null;
                    cell.setCellValue(s);
                }
                col++;
                cell.setCellStyle(contentStyle(workbook, row));
            }
            for (; col < 5; col++) {
                XSSFCell cell = dataRow.createCell(col);
                cell.setCellValue("");
                cell.setCellStyle(contentStyle(workbook, row));
            }
        }

        // 数据列去重
        colList = colList.stream().

                distinct().

                collect(Collectors.toList());

        int lastRowNum = sheet.getLastRowNum();


        // 统计数据列总数据
        CellRangeAddress tableFootCell = null;
        if (excelEntity.getStatistics()) {
            XSSFRow totalRow = sheet.createRow(lastRowNum + 1);
            XSSFCell totalRowCell = totalRow.createCell(0);
            totalRowCell.setCellValue("列总数");
            totalRowCell.setCellStyle(contentStyle(workbook, lastRowNum + 1));
            // 构建统计函数
            int num = 26;
            String formulaTotalEnd = null;
            for (int outCol = 0; outCol <= colList.get(colList.size()-1) / num; outCol++) {
                int innerCol = colList.size() - outCol * num >= num ? 26 : colList.size() - outCol * num;
                for (int col = 0; col < innerCol; col++) {
                    int startCol = 65;
                    XSSFCell cell = totalRow.createCell(colList.get(outCol * num + col));
                    startCol += colList.get(outCol * num + col);
                    String formula;
                    if (startCol >= num + 65) {
                        startCol = startCol - outCol * num;
                        if (startCol == num + 65) {
                            startCol -= num;
                            if (excelEntity.isStatisticsTotal()) {
                                formula = "IMSUB(" + ((char) (outCol + 65)) + ((char) (startCol)) + ","
                                        + ((char) (outCol + 65)) + ((char) (startCol)) + ")";
                            } else {
                                formula = "sum(" + ((char) (outCol + 65)) + ((char) (startCol)) + (realRow + 1) + ":"
                                        + ((char) (outCol + 65)) + ((char) (startCol)) + (lastRowNum + 1) + ")";
                            }
                            if (excelEntity.isStatisticsTotalAll()) {
                                formulaTotalEnd = ":" + ((char) (outCol + 64)) + ((char) (startCol)) + (lastRowNum + 2) + ")";
                            }
                        } else {
                            if (excelEntity.isStatisticsTotal()) {
                                formula = "IMSUB(" + ((char) (outCol + 64)) + ((char) (startCol)) + ","
                                        + ((char) (outCol + 64)) + ((char) (startCol)) + ")";
                            } else {
                                formula = "sum(" + ((char) (outCol + 64)) + ((char) (startCol)) + (realRow + 1) + ":"
                                        + ((char) (outCol + 64)) + ((char) (startCol)) + (lastRowNum + 1) + ")";
                            }
                            if (excelEntity.isStatisticsTotalAll()) {
                                formulaTotalEnd = ":" + ((char) (outCol + 64)) + ((char) (startCol)) + (lastRowNum + 2) + ")";
                            }
                        }
                    } else {
                        if (excelEntity.isStatisticsTotal()) {
                            formula = "IMSUB(" + ((char) startCol) + "," + ((char) startCol) + ")";
                        } else {
                            formula = "sum(" + ((char) startCol) + (realRow + 1) + ":" + ((char) startCol) + (lastRowNum + 1) + ")";
                        }
                        if (excelEntity.isStatisticsTotalAll()) {
                            formulaTotalEnd = ":" + ((char) (startCol)) + (lastRowNum + 2) + ")";
                        }
                    }
                    if (excelEntity.isStatisticsTotal()) {
                        String[] formulaList = formula.split(",");
                        String[] split = formulaNumList.get(col).split("-");
                        if (split.length > 0) {
                            for (int i = 0; i < split.length; i++) {
                                int result = realRow + 1 + Integer.parseInt(split[i]);
                                split[i] = result + "";
                            }
                            cell.setCellFormula("ROUND(" + formulaList[0] + split[0] + "," +
                                    formulaList[1].replace(")", "") + split[1] + "),2)");
                        }
                    } else {
                        cell.setCellFormula("ROUND(" + formula + ",2)");
                    }
                    cell.setCellStyle(contentStyle(workbook, lastRowNum + 1));
                }
            }
            // 补充统计列数
            for (int col = colList.get(colList.size() - 1) + 1; col < 5; col++) {
                XSSFCell cell = totalRow.createCell(col);
                cell.setCellValue("");
                cell.setCellStyle(contentStyle(workbook, lastRowNum + 1));
            }


            if (excelEntity.isStatisticsTotalAll()) {
                String formulaTotalStart = "sum(" + (char) (colList.get(0) + 65) + (lastRowNum + 2);
                String formula = formulaTotalStart + formulaTotalEnd;
                XSSFRow allTotalRow = sheet.createRow(lastRowNum + 2);
                XSSFCell cell = allTotalRow.createCell(0);
                cell.setCellValue("全总数");
                cell.setCellStyle(contentStyle(workbook, lastRowNum + 2));
                XSSFCell cell1 = allTotalRow.createCell(1);
                // 合并列数
                tableFootCell = new CellRangeAddress(lastRowNum + 2, lastRowNum + 2, 1, tableFiled.size() - 1 > 4 ? tableFiled.size() - 1 : 4);
                sheet.addMergedRegion(tableFootCell);
                cell1.setCellFormula("ROUND(" + formula + ",2)");
                CellStyle cellStyle = contentStyle(workbook, lastRowNum + 2);
                cellStyle.setAlignment(HorizontalAlignment.LEFT);
                cell1.setCellStyle(cellStyle);
            }

        }

        // 设置边框
        defaultBorderSetting(workbook, sheet, tableHeadCell, tableFootCell, lastRowNum, 5);

        // 生成统计图
        if (excelEntity.getExcelChartTypeEnum() != null) {
            switch (excelEntity.getExcelChartTypeEnum()) {
                case LINE_CHAR:
                    tableFiled.remove(tableFiled.keySet().iterator().next());
                    ExcelImgUtil.createSheetLineChart(workbook.getSheet(excelEntity.getTableName()), excelEntity.getTableName(),
                            tableFiled.values().stream().collect(Collectors.toList()), null != filedUnit ? filedUnit.get("x").toString() : "",
                            null != filedUnit ? filedUnit.get("y").toString() : "", realRow, tableDataList.get(filedSet.iterator().next()).size(),
                            0, 0, colList, null);
                    break;
                case PIE:
                    ExcelImgUtil.createSheetPieChart(workbook.getSheet(excelEntity.getTableName()), excelEntity.getTableName(),
                            realRow, tableDataList.get(filedSet.iterator().next()).size() + 1, 0, 0,
                            colList, isStatistics, null);
                    break;
                case BAR_CHAR:
                    ExcelImgUtil.createSheetBarChart(workbook.getSheet(excelEntity.getTableName()), excelEntity.getTableName(),
                            null != filedUnit ? filedUnit.get("x").toString() : "", null != filedUnit ? filedUnit.get("y").toString() : "",
                            realRow, tableDataList.get(filedSet.iterator().next()).size() + 1, 0, 0, colList, isStatistics, null);
                    break;
                case LINE_PIE_BAR_CHAR:
                    tableFiled.remove(tableFiled.keySet().iterator().next());
                    XSSFChart chart = ExcelImgUtil.getSheetDrawing(workbook.getSheet(excelEntity.getTableName()), "数据分析", 0, 2, 17, 0, 7);
                    ExcelImgUtil.createSheetLineChart(workbook.getSheet(excelEntity.getTableName()), "数据分析",
                            tableFiled.values().stream().collect(Collectors.toList()), null != filedUnit ? filedUnit.get("x").toString() : "",
                            null != filedUnit ? filedUnit.get("y").toString() : "", realRow, tableDataList.get(filedSet.iterator().next()).size(),
                            0, 0, colList, chart);
                    chart = ExcelImgUtil.getSheetDrawing(workbook.getSheet(excelEntity.getTableName()), "用量分析", 0, 20, 37, 0, 7);
                    ExcelImgUtil.createSheetPieChart(workbook.getSheet(excelEntity.getTableName()), "用量分析",
                            realRow, tableDataList.get(filedSet.iterator().next()).size() + 1, 0, 0,
                            colList, isStatistics, chart);
                    chart = ExcelImgUtil.getSheetDrawing(workbook.getSheet(excelEntity.getTableName()), "用量占比", 0, 40, 54, 0, 7);
                    ExcelImgUtil.createSheetBarChart(workbook.getSheet(excelEntity.getTableName()), "用量占比",
                            null != filedUnit ? filedUnit.get("x").toString() : "", null != filedUnit ? filedUnit.get("y").toString() : "",
                            realRow, tableDataList.get(filedSet.iterator().next()).size() + 1, 0, 0, colList, isStatistics, chart);
                default:
                    break;
            }
        }

        // 配置默认打印设置
        defaultPrintSetting(sheet);

    }

    /**
     * 默认边界设置
     *
     * @param sheet         工作簿sheet
     * @param rowNum        行数
     * @param colNum        列数
     * @param workbook      工作簿
     * @param tableHeadCell 表头部
     * @param tableFootCell 表足
     */
    private static void defaultBorderSetting(XSSFWorkbook workbook, XSSFSheet sheet, CellRangeAddress tableHeadCell, CellRangeAddress tableFootCell, int rowNum, Integer colNum) {

        // 第一行设置上边框
        RegionUtil.setBorderTop(BorderStyle.THIN, tableHeadCell, sheet);
        RegionUtil.setBorderRight(BorderStyle.THIN, tableHeadCell, sheet);
        RegionUtil.setBorderLeft(BorderStyle.THIN, tableHeadCell, sheet);
        RegionUtil.setBorderBottom(BorderStyle.THIN, tableHeadCell, sheet);
        RegionUtil.setBottomBorderColor(IndexedColors.GREY_50_PERCENT.getIndex(), tableHeadCell, sheet);
        RegionUtil.setLeftBorderColor(IndexedColors.GREY_50_PERCENT.getIndex(), tableHeadCell, sheet);
        RegionUtil.setRightBorderColor(IndexedColors.GREY_50_PERCENT.getIndex(), tableHeadCell, sheet);
        RegionUtil.setTopBorderColor(IndexedColors.GREY_50_PERCENT.getIndex(), tableHeadCell, sheet);

        // 最后一行设置下边框
        if (null != tableFootCell) {
            RegionUtil.setBorderTop(BorderStyle.THIN, tableFootCell, sheet);
            RegionUtil.setBorderRight(BorderStyle.THIN, tableFootCell, sheet);
            RegionUtil.setBorderLeft(BorderStyle.THIN, tableFootCell, sheet);
            RegionUtil.setBorderBottom(BorderStyle.THIN, tableFootCell, sheet);
            RegionUtil.setBottomBorderColor(IndexedColors.GREY_50_PERCENT.getIndex(), tableFootCell, sheet);
            RegionUtil.setLeftBorderColor(IndexedColors.GREY_50_PERCENT.getIndex(), tableFootCell, sheet);
            RegionUtil.setRightBorderColor(IndexedColors.GREY_50_PERCENT.getIndex(), tableFootCell, sheet);
            RegionUtil.setTopBorderColor(IndexedColors.GREY_50_PERCENT.getIndex(), tableFootCell, sheet);
        } else {
            XSSFRow lastRow = sheet.getRow(rowNum);
            XSSFCell cell = lastRow.getCell(0);
            if (null == cell) {
                cell = lastRow.createCell(0);
            }
            XSSFCellStyle cellStyle = cell.getCellStyle();
            if (null == cellStyle) {
                cellStyle = workbook.createCellStyle();
            }
            cellStyle.setBorderBottom(BorderStyle.THIN);
            cellStyle.setBorderLeft(BorderStyle.THIN);
            cellStyle.setBorderRight(BorderStyle.THIN);
            cellStyle.setBorderTop(BorderStyle.THIN);
            cell.setCellStyle(cellStyle);


        }
    }

    /**
     * 配置默认打印设置
     *
     * @param sheet 表
     */
    private static void defaultPrintSetting(XSSFSheet sheet) {
        XSSFPrintSetup printSetup = sheet.getPrintSetup();
        // 横纵向设置（纵向）
        printSetup.setLandscape(false);

        // 设置纸张（A3）
        printSetup.setPaperSize(PrintSetup.A3_PAPERSIZE);

        // 设置将工作表调整为一页显示
        sheet.setFitToPage(true);

        // 页码设置
        Footer footer = sheet.getFooter();
        footer.setCenter("第" + HSSFFooter.page() + "页，共" + HSSFFooter.numPages() + "页");

    }


    /**
     * 反射来填充excelEntity中的数据
     *
     * @param excelEntity
     * @param constantMap
     * @param objects
     */
    public static void fillExcelEntity(ExcelEntity excelEntity, Map<String, Map<String, String>> constantMap, List<?> objects, String messageId) {
        Map<String, List<Object>> tableDataList = new HashMap<>(16);
        LinkedHashMap<String, String> tableFiled = excelEntity.getTableFiled();
        Set<String> filedSet = tableFiled.keySet();
        for (String filed : filedSet) {
            List<Object> data = new ArrayList<>();
            for (Object object : objects) {
                String methodName = "get" + filed.replaceFirst(filed.substring(0, 1), filed.substring(0, 1).toUpperCase());
                Object value = null;
                try {
                    Method method = object.getClass().getMethod(methodName);
                    value = method.invoke(object);

                    if (null != value) {
                        // boolean值转换
                        if ("true".equals(value.toString()) || "false".equals(value.toString())) {
                            value = constantMap.get(filed).get(value.toString());
                        }
                        // valueList值处理
                        if (value instanceof Map) {
                            Map<String, String> valueMap = JSONObject.parseObject(JSONObject.toJSONString(value), new TypeReference<Map<String, String>>() {
                            });
                            LinkedHashMap<String, String> valueList = JSONObject.parseObject(tableFiled.get(filed), new TypeReference<LinkedHashMap<String, String>>() {
                            });
                            Set<String> values = valueList.keySet();
                            values.forEach(valueInfo -> {
                                List<Object> dataMap = tableDataList.get(valueInfo);
                                if (null != dataMap) {
                                    dataMap.add(null == valueMap.get(valueInfo) ? "--" : dateToStringIfIsDate(valueMap.get(valueInfo)));
                                } else {
                                    dataMap = new ArrayList<>();
                                    dataMap.add(null == valueMap.get(valueInfo) ? "--" : dateToStringIfIsDate(valueMap.get(valueInfo)));
                                }
                                tableDataList.put(valueInfo, dataMap);
                                tableFiled.put(valueInfo, valueInfo);
                            });
                            continue;
                        } else {
                            data.add(dateToStringIfIsDate(value));
                        }
                    } else {
                        data.add("--");
                    }
                } catch (NoSuchMethodException e) {
                } catch (InvocationTargetException e) {
                    throw new RuntimeException(e);
                } catch (IllegalAccessException e) {
                    throw new RuntimeException(e);
                }

            }
            if (!"valueList".equals(filed)) {
                tableDataList.put(filed, data);
            }
        }
        tableFiled.remove("valueList");
        excelEntity.setTableFiled(tableFiled);
        excelEntity.setTableDataList(tableDataList);
    }

    /**
     * 创建内容行样式
     *
     * @param wb 工作簿
     * @return
     */
    public static CellStyle contentStyle(Workbook wb, int row) {
        CellStyle contentStyle = wb.createCellStyle();
        Font contentFont = wb.createFont();
        contentFont.setFontName("宋体");
        contentFont.setColor(HSSFFont.COLOR_NORMAL);
        contentFont.setFontHeightInPoints((short) 12);


        contentStyle.setAlignment(HorizontalAlignment.CENTER);
        contentStyle.setFont(contentFont);
        // 设置边框以及边框颜色
        contentStyle.setBorderBottom(BorderStyle.THIN);
        contentStyle.setBorderLeft(BorderStyle.THIN);
        contentStyle.setBorderRight(BorderStyle.THIN);
        contentStyle.setBorderTop(BorderStyle.THIN);
        contentStyle.setBottomBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        contentStyle.setLeftBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        contentStyle.setRightBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        contentStyle.setTopBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());

        // 奇数行数填充颜色
        if (row % 2 != 0) {
            contentStyle.setFillForegroundColor(IndexedColors.PALE_BLUE.getIndex());
        } else {
            // 偶数行数填充颜色
            contentStyle.setFillForegroundColor(IndexedColors.WHITE.getIndex());
        }
        contentStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return contentStyle;
    }


    /**
     * 创建标题行样式
     *
     * @param wb
     * @return
     */
    public static CellStyle headStyle(Workbook wb, boolean isWrapText, Boolean isRead) {
        // 创建样式对象
        CellStyle headStyle = wb.createCellStyle();
        // 创建字体
        Font headFont = wb.createFont();
        headFont.setFontName("宋体");
        headFont.setBold(true);
        headFont.setColor(HSSFFont.COLOR_NORMAL);
        headFont.setFontHeightInPoints((short) 11);
        headStyle.setAlignment(HorizontalAlignment.CENTER);
        headStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        headFont.setColor(IndexedColors.WHITE.getIndex());
        if (isRead) {
            headFont.setColor(IndexedColors.RED.getIndex());
        }
        headStyle.setFont(headFont);

        headStyle.setBorderBottom(BorderStyle.THIN);
        headStyle.setBorderLeft(BorderStyle.THIN);
        headStyle.setBorderRight(BorderStyle.THIN);
        headStyle.setBorderTop(BorderStyle.THIN);
        headStyle.setTopBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        headStyle.setRightBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        headStyle.setBottomBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        headStyle.setLeftBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());

        headStyle.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
        headStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        // 设置自动换行
        headStyle.setWrapText(isWrapText);
        return headStyle;
    }


    private static Object dateToStringIfIsDate(Object value) {
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss", Locale.CHINA);
        if (value instanceof Date) {
            return simpleDateFormat.format(value);
        }
        if (value instanceof LocalDateTime) {
            return dateTimeFormatter.format((LocalDateTime) value);
        }
        if (value instanceof LocalDate) {
            return dateFormatter.format((LocalDate) value);
        }
        if (value instanceof LocalTime) {
            return timeFormatter.format((LocalTime) value);
        } else {
            return value;
        }
    }
}
