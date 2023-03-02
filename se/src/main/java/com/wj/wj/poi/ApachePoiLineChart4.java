package com.wj.wj.poi;


import com.wj.wj.utils.ExcelUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Created by IntelliJ IDEA
 *
 * @author wj
 * @date 2022/10/26 13:46
 */
public class ApachePoiLineChart4 {

    public static void main(String[] args) throws IOException {
        XSSFWorkbook wb = new XSSFWorkbook();
        String sheetName = "智能水表";
        FileOutputStream fileOut = null;
        List<Map<String,Object>> data = new ArrayList<>();
        Map<String, Object> map = new HashMap<>();
        map.put("date","2022-10-26 13:00:00");
        map.put("value",3399.92);
        data.add(map);
        map = new HashMap<>();
        map.put("date","2022-10-26 12:00:00");
        map.put("value",3397);
        data.add(map);
        map = new HashMap<>();
        map.put("date","2022-10-26 11:00:00");
        map.put("value",3387.59);
        data.add(map);
        map = new HashMap<>();
        map.put("date","2022-10-26 10:00:00");
        map.put("value",3395.96);
        data.add(map);
        map = new HashMap<>();
        map.put("date","2022-10-26 09:00:00");
        map.put("value",3376.87);
        data.add(map);
        map = new HashMap<>();
        map.put("date","2022-10-26 08:00:00");
        map.put("value",3395.7);
        data.add(map);
        map = new HashMap<>();
        map.put("date","2022-10-26 07:00:00");
        map.put("value",3398.31);
        data.add(map);
        map = new HashMap<>();
        map.put("date","2022-10-26 06:00:00");
        map.put("value",3388.98);
        data.add(map);
        map = new HashMap<>();
        map.put("date","2022-10-26 05:00:00");
        map.put("value",3389.62);
        data.add(map);
        map = new HashMap<>();
        map.put("date","2022-10-26 04:00:00");
        map.put("value",3396.62);
        data.add(map);
        map = new HashMap<>();
        map.put("date","2022-10-26 03:00:00");
        map.put("value",3396.15);
        data.add(map);
        map = new HashMap<>();
        map.put("date","2022-10-26 02:00:00");
        map.put("value",3393.27);
        data.add(map);
        map = new HashMap<>();
        map.put("date","2022-10-26 01:00:00");
        map.put("value",3392.82);
        data.add(map);

        try {
            XSSFSheet sheet = wb.createSheet(sheetName);

            //第一行，表头
            Row row = sheet.createRow(0);
            Cell cell = row.createCell(0);
            cell.setCellValue(sheetName);
            sheet.addMergedRegion(new CellRangeAddress(0,0,0,1));

            row = sheet.createRow(1);
            cell = row.createCell(0);
            cell.setCellValue("时间");
            cell = row.createCell(1);
            cell.setCellValue("水能");

            row = sheet.createRow(2);
            cell = row.createCell(0);
            cell.setCellValue("2022-10-26 13:00:00");
            cell = row.createCell(1);
            cell.setCellValue(3399.92);

            row = sheet.createRow(3);
            cell = row.createCell(0);
            cell.setCellValue("2022-10-26 12:00:00");
            cell = row.createCell(1);
            cell.setCellValue(3397);

            row = sheet.createRow(4);
            cell = row.createCell(0);
            cell.setCellValue("2022-10-26 11:00:00");
            cell = row.createCell(1);
            cell.setCellValue(3397);

            row = sheet.createRow(5);
            cell = row.createCell(0);
            cell.setCellValue("2022-10-26 10:00:00");
            cell = row.createCell(1);
            cell.setCellValue(3365);

            row = sheet.createRow(6);
            cell = row.createCell(0);
            cell.setCellValue("2022-10-26 09:00:00");
            cell = row.createCell(1);
            cell.setCellValue(2356);

            row = sheet.createRow(7);
            cell = row.createCell(0);
            cell.setCellValue("2022-10-26 08:00:00");
            cell = row.createCell(1);
            cell.setCellValue(3546);

            row = sheet.createRow(8);
            cell = row.createCell(0);
            cell.setCellValue("2022-10-26 07:00:00");
            cell = row.createCell(1);
            cell.setCellValue(3245);

            row = sheet.createRow(9);
            cell = row.createCell(0);
            cell.setCellValue("2022-10-26 06:00:00");
            cell = row.createCell(1);
            cell.setCellValue(3397);

            row = sheet.createRow(10);
            cell = row.createCell(0);
            cell.setCellValue("2022-10-26 05:00:00");
            cell = row.createCell(1);
            cell.setCellValue(3389.62);

            row = sheet.createRow(11);
            cell = row.createCell(0);
            cell.setCellValue("2022-10-26 04:00:00");
            cell = row.createCell(1);
            cell.setCellValue(3367.53);

            row = sheet.createRow(12);
            cell = row.createCell(0);
            cell.setCellValue("2022-10-26 03:00:00");
            cell = row.createCell(1);
            cell.setCellValue(3397);

            row = sheet.createRow(13);
            cell = row.createCell(0);
            cell.setCellValue("2022-10-26 02:00:00");
            cell = row.createCell(1);
            cell.setCellValue(3393.27);

            row = sheet.createRow(14);
            cell = row.createCell(0);
            cell.setCellValue("2022-10-26 01:00:00");
            cell = row.createCell(1);
            cell.setCellValue(3392.82);

            // 第二行以后，数据

//            for(int i=0;i< data.size();i++){
//                int count = i+1;
//                Row rowData = sheet.createRow(count);
//                Cell cellDate = rowData.createCell(0);
//                cellDate.setCellValue((String) data.get(i).get("date"));
//                cellDate = rowData.createCell(1);
//                cellDate.setCellValue((data.get(i).get("value").toString()));
//            }



//            ExcelUtils.createLineChart(sheet,yList,"智能水表历史数据",0,14,0,1,0,0,4,16);
            ExcelUtils.createSheetLineChark(sheet,"智能水表历史数据","智能水表","时间","水能",2, 14,0,0,1);

//            //创建一个画布
//            XSSFDrawing drawing = sheet.createDrawingPatriarch();
//            //前四个默认0，[0,5]：从0列5行开始;[7,26]:宽度7个单元格，26向下扩展到26行
//            //默认宽度(14-8)*12
//            XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 0, 14, 7, 26);
//            //创建一个chart对象
//            XSSFChart chart = drawing.createChart(anchor);
//            //标题
//            chart.setTitleText("智能水表历史数据");
//            //标题覆盖
//            chart.setTitleOverlay(false);
//
//            //图例位置
//            XDDFChartLegend legend = chart.getOrAddLegend();
//            legend.setPosition(LegendPosition.TOP);
//
//            //分类轴标(X轴),标题位置
//            XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
//            bottomAxis.setTitle("时间");
//            //值(Y轴)轴,标题位置
//            XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
//            leftAxis.setTitle("用水量：m^3");
//
//            //CellRangeAddress(起始行号，终止行号， 起始列号，终止列号）
//            //分类轴标(X轴)数据，单元格范围位置[1, 0]到[13, 0]
//            XDDFDataSource<String> countries = XDDFDataSourcesFactory.fromStringCellRange(sheet, new CellRangeAddress(1, 12, 0, 0));
//
//            //水表数据，单元格范围位置[1, 1]到[13, 1]
//            XDDFNumericalDataSource<Double> area = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(1, 12, 1, 1));
//
//            //LINE：折线图，
//            XDDFLineChartData lineCart = (XDDFLineChartData) chart.createData(ChartTypes.LINE, bottomAxis, leftAxis);
//
//            //图表加载数据，折线1
//            XDDFLineChartData.Series series1 = (XDDFLineChartData.Series) lineCart.addSeries(countries, area);
//            //折线图例标题
//            series1.setTitle("面积", null);
//            //直线
//            series1.setSmooth(false);
//            //设置标记大小
//            series1.setMarkerSize((short) 6);
//            //设置标记样式，星星
//            series1.setMarkerStyle(MarkerStyle.STAR);
//
////            绘制
//            chart.plot(lineCart);

            // 将输出写入excel文件
            String filename = "智能水表历史数据.xlsx";
            fileOut = new FileOutputStream(filename);
            wb.write(fileOut);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            wb.close();
            if (fileOut != null) {
                fileOut.close();
            }
        }

    }
}
