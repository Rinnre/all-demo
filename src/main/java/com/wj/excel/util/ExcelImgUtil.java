package com.wj.excel.util;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPlotArea;

import java.util.ArrayList;
import java.util.List;

/**
 * Created by IntelliJ IDEA
 *
 * @author wj
 * @date 2022/10/27 10:28
 */
public class ExcelImgUtil {


    /**
     * 柱状图
     *
     * @param sheet
     * @param chartTitle
     * @param firstRow
     * @param dataEndRow
     * @param startCol
     * @param xAxisCol
     * @param valueCol
     */
    public static void createSheetBarChart(XSSFSheet sheet, String chartTitle, String xName, String yName, int firstRow, int dataEndRow, int startCol, int xAxisCol, List<Integer> valueCol, boolean isStatistics) {
        //根据列数创建一个画布
        int maxCol = 9;
        int charWidth = 6;
        int charHeight = 10;
        XSSFChart chart;
        if (maxCol - valueCol.size() < charWidth) {
            // 在表格下方创建画布
            chart = getSheetDrawing(sheet, chartTitle, startCol, 2, dataEndRow>charHeight?dataEndRow:charHeight, 0, charWidth);
        } else {
            // 在表格右边创建画布
            chart = getSheetDrawing(sheet, chartTitle, startCol, -sheet.getLastRowNum() + 1, dataEndRow>charHeight?0:charHeight, valueCol.get(valueCol.size() - 1) + 2, valueCol.size() + 2 + charWidth);
        }


        //标题
        chart.setTitleText(chartTitle);
        //标题覆盖
        chart.setTitleOverlay(false);

        //图例位置
        XDDFChartLegend legend = chart.getOrAddLegend();
        legend.setPosition(LegendPosition.TOP_RIGHT);

        //分类轴标(X轴),标题位置
        XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
        bottomAxis.setTitle("设备");
        //值(Y轴)轴,标题位置
        XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
        leftAxis.setTitle(yName);

        //CellRangeAddress(起始行号，终止行号， 起始列号，终止列号）
        //分类轴标数据，
        XDDFDataSource<String> xData;
        XDDFNumericalDataSource<Double> values;
        if (isStatistics) {
            xData = XDDFDataSourcesFactory.fromStringCellRange(sheet, new CellRangeAddress(firstRow - 1, firstRow - 1, valueCol.get(0), valueCol.get(valueCol.size() - 1)));
            //数据1，单元格范围位置[1, 0]到[1, 6]
            values = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(dataEndRow + 1, dataEndRow + 1, valueCol.get(0), valueCol.get(valueCol.size() - 1)));
        } else {
            xData = XDDFDataSourcesFactory.fromStringCellRange(sheet, new CellRangeAddress(firstRow , dataEndRow, 0, 0));
            //数据1，单元格范围位置[1, 0]到[1, 6]
            values = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(firstRow , dataEndRow , valueCol.get((valueCol.size() - 1)), valueCol.get(valueCol.size() - 1)));
        }

        //bar：条形图，
        XDDFBarChartData bar = (XDDFBarChartData) chart.createData(ChartTypes.BAR, bottomAxis, leftAxis);

        leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);
        leftAxis.setCrossBetween(AxisCrossBetween.BETWEEN);
        //设置为可变颜色
        bar.setVaryColors(true);
        //条形图方向，纵向/横向：纵向
        bar.setBarDirection(BarDirection.COL);

        //图表加载数据，条形图1
        XDDFBarChartData.Series series1 = (XDDFBarChartData.Series) bar.addSeries(xData, values);

        //条形图例标题
        series1.setTitle("能耗", null);

        CTPlotArea plotArea = chart.getCTChart().getPlotArea();
        //绘制
        chart.plot(bar);

    }


    /**
     * 饼图
     *
     * @param sheet      sheet
     * @param chartTitle 标题
     * @param firstRow   开始行
     * @param dataEndRow 结束行
     * @param startCol   开始列
     * @param xAxisCol   x轴列
     * @param valueCol   结束列
     */
    public static void createSheetPieChart(XSSFSheet sheet, String chartTitle, int firstRow, int dataEndRow, int startCol, int xAxisCol, List<Integer> valueCol, boolean isStatistics) {
        //根据列数创建一个画布
        int maxCol = 9;
        int charWidth = 6;
        int charHeight = 10;
        XSSFChart chart;
        if (maxCol - valueCol.size() < charWidth) {
            // 在表格下方创建画布
            chart = getSheetDrawing(sheet, chartTitle, startCol, 2, dataEndRow>charHeight?dataEndRow:charHeight, 0, charWidth);
        } else {
            // 在表格右边创建画布
            chart = getSheetDrawing(sheet, chartTitle, startCol, -sheet.getLastRowNum() + 1, dataEndRow>charHeight?0:charHeight, valueCol.get(valueCol.size() - 1) + 2, valueCol.size() + 2 + charWidth);
        }


        //标题
        chart.setTitleText(chartTitle);
        //标题是否覆盖图表
        chart.setTitleOverlay(false);
        //图例位置
        XDDFChartLegend legend = chart.getOrAddLegend();
        legend.setPosition(LegendPosition.TOP_RIGHT);


        //CellRangeAddress(起始行号，终止行号， 起始列号，终止列号）
        //分类轴标数据，
        XDDFDataSource<String> xData;
        XDDFNumericalDataSource<Double> values;
        if (isStatistics) {
            xData = XDDFDataSourcesFactory.fromStringCellRange(sheet, new CellRangeAddress(firstRow - 1, firstRow - 1, valueCol.get(0), valueCol.get(valueCol.size() - 1)));
            //数据1，单元格范围位置[1, 0]到[1, 6]
            values = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(dataEndRow + 1, dataEndRow + 1, valueCol.get(0), valueCol.get(valueCol.size() - 1)));
        } else {
            xData = XDDFDataSourcesFactory.fromStringCellRange(sheet, new CellRangeAddress(firstRow , dataEndRow, 0, 0));
            //数据1，单元格范围位置[1, 0]到[1, 6]
            values = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(firstRow , dataEndRow , valueCol.get((valueCol.size() - 1)), valueCol.get(valueCol.size() - 1)));
        }

        //分类轴标(X轴),标题位置
        XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
        //值(Y轴)轴,标题位置
        XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
        XDDFChartData data = chart.createData(ChartTypes.PIE, bottomAxis, leftAxis);

        //设置为可变颜色
        data.setVaryColors(true);
        //图表加载数据
        data.addSeries(xData, values);

        //绘制
        chart.plot(data);

    }


    /**
     * @param sheet
     * @param chartTitle
     * @param lineName
     * @param xName
     * @param yName
     * @param firstRow
     * @param dataEndRow
     * @param startCol
     * @param xAxisCol   x轴
     * @param valueCol   条数
     */
    public static void createSheetLineChart(XSSFSheet sheet, String chartTitle, List<String> lineName, String xName, String yName, int firstRow, int dataEndRow, int startCol, int xAxisCol, List<Integer> valueCol) {
        //根据列数创建一个画布
        int maxCol = 9;
        int charWidth = 6;
        int charHeight = 10;
        XSSFChart chart;
        if (maxCol - valueCol.size() < charWidth) {
            // 在表格下方创建画布
            chart = getSheetDrawing(sheet, chartTitle, startCol, 2, dataEndRow>charHeight?dataEndRow:charHeight, 0, charWidth);
        } else {
            // 在表格右边创建画布
            chart = getSheetDrawing(sheet, chartTitle, startCol, -sheet.getLastRowNum() + 1, dataEndRow>charHeight?0:charHeight, valueCol.get(valueCol.size() - 1) + 2, valueCol.size() + 2 + charWidth);
        }

        //图例设置
        if (valueCol.size() > 1) {
            XDDFChartLegend legend = chart.getOrAddLegend();
            legend.setPosition(LegendPosition.TOP_RIGHT);
        }

        // 分类轴标(X轴),标题位置


        XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
        if (StringUtils.isNotBlank(xName)) {
            bottomAxis.setTitle(xName);
        }

        //值(Y轴)轴,标题位置
        XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
        if (StringUtils.isNotBlank(yName))
            leftAxis.setTitle(yName);
        //CellRangeAddress(起始行号，终止行号， 起始列号，终止列号）

        XDDFDataSource<String> xData = XDDFDataSourcesFactory.fromStringCellRange(sheet, new CellRangeAddress(firstRow, dataEndRow + 1, xAxisCol, xAxisCol));

        //数据1，单元格范围位置[1, 0]到[1, 6]
        List<XDDFNumericalDataSource<Double>> lines = new ArrayList<>();
        for (Integer num : valueCol) {
            XDDFNumericalDataSource<Double> values = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(firstRow, dataEndRow + 1, num, num));
            lines.add(values);
        }

        //LINE：折线图，
        XDDFLineChartData data = (XDDFLineChartData) chart.createData(ChartTypes.LINE, bottomAxis, leftAxis);

        //图表加载数据，折线
        for (int num = 0; num < lines.size(); num++) {

            XDDFNumericalDataSource<Double> line = lines.get(num);
            XDDFLineChartData.Series series = (XDDFLineChartData.Series) data.addSeries(xData, line);
            //折线图例标题
            series.setTitle(lineName.get(num), null);

            //直线
            series.setSmooth(false);
            series.setMarkerStyle(MarkerStyle.NONE);

        }


        //绘制
        chart.plot(data);
    }

    /**
     * @param sheet          sheet表格
     * @param chartTitle     图形标题
     * @param startCol       图表开始列数
     * @param firstRowOffset 与数据表格的偏移函数
     * @param endRowOffset   结束行数偏移数
     * @param startColOffset 开始列数偏移数
     * @param endColOffset   结束列数偏移数
     * @return chart对象
     */
    private static XSSFChart getSheetDrawing(XSSFSheet sheet, String chartTitle, int startCol, int firstRowOffset, int endRowOffset, int startColOffset, int endColOffset) {
        //创建一个画布
        XSSFDrawing drawing = sheet.createDrawingPatriarch();

        //默认宽度(14-8)*12
        int lastRowNum = sheet.getLastRowNum();
        XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, startCol + startColOffset, lastRowNum + firstRowOffset, startCol + endColOffset, lastRowNum + endRowOffset);
        //创建一个chart对象
        XSSFChart chart = drawing.createChart(anchor);
        //标题
        chart.setTitleText(chartTitle);
        //标题是否覆盖图表
        chart.setTitleOverlay(false);
        return chart;
    }
}
