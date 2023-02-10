package com.wj.wj.poi;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.*;
import org.openxmlformats.schemas.drawingml.x2006.chart.*;

import java.io.FileOutputStream;
import java.io.IOException;

/**
 * Created by IntelliJ IDEA
 *
 * @author wj
 * @date 2022/10/28 9:47
 */
public class LineChart {
    public static void main(String[] args) throws IOException {
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet("Sheet1");

        // 数据模拟
        Row row;
        Cell cell;

        row = sheet.createRow(0);
        row.createCell(0);
        row.createCell(1).setCellValue("Lines");

        for (int r = 1; r < 7; r++) {
            row = sheet.createRow(r);
            cell = row.createCell(0);
            cell.setCellValue("C" + r);
            cell = row.createCell(1);
            cell.setCellValue(new java.util.Random().nextDouble() * 10d);
        }
        // 生成图片
        //创建一个画布
        XSSFDrawing drawing = sheet.createDrawingPatriarch();

        XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 4, 0, 11, 15);
        //创建一个chart对象
        XSSFChart chart = drawing.createChart(anchor);

        CTChart ctChart = ((XSSFChart)chart).getCTChart();
        CTPlotArea ctPlotArea = ctChart.getPlotArea();

        // 设置标题
        ((XSSFChart) chart).setTitleText("图表Demo");
        // 标题的位置（于图表上方或居中覆盖）
        ctChart.getTitle().addNewOverlay().setVal(false);
        ctChart.addNewShowDLblsOverMax().setVal(true);
        CTTitle tt  = ctChart.addNewTitle();
        ctChart.setTitle(tt);

        CTLineChart ctLineChart = ctPlotArea.addNewLineChart();
        CTBoolean ctBoolean = ctLineChart.addNewVaryColors();
        ctBoolean.setVal(true);

        //the line series
        CTLineSer ctLineSer = ctLineChart.addNewSer();
        CTSerTx ctSerTx = ctLineSer.addNewTx();
        CTStrRef ctStrRef = ctSerTx.addNewStrRef();
        ctStrRef.setF("Sheet1!$C$1");
        ctLineSer.addNewIdx().setVal(1);
        CTAxDataSource cttAxDataSource = ctLineSer.addNewCat();
        ctStrRef = cttAxDataSource.addNewStrRef();
        ctStrRef.setF("Sheet1!$A$2:$A$7");
        CTNumDataSource ctNumDataSource = ctLineSer.addNewVal();
        CTNumRef ctNumRef = ctNumDataSource.addNewNumRef();
        ctNumRef.setF("Sheet1!$C$2:$C$7");

        ctLineChart.addNewAxId().setVal(123458); //cat axis 2 (lines)
        ctLineChart.addNewAxId().setVal(123459); //val axis 2 (right)

        CTCatAx ctCatAx = ctPlotArea.addNewCatAx();
        ctCatAx.addNewAxId().setVal(123458); //id of the cat axis
        CTScaling ctScaling = ctCatAx.addNewScaling();
        ctScaling.addNewOrientation().setVal(STOrientation.MIN_MAX);
        ctCatAx.addNewDelete().setVal(true); //this cat axis is deleted
        ctCatAx.addNewAxPos().setVal(STAxPos.B);
        ctCatAx.addNewCrossAx().setVal(123459); //id of the val axis
        ctCatAx.addNewTickLblPos().setVal(STTickLblPos.NEXT_TO);


         // 导出文件
        FileOutputStream fileOut = new FileOutputStream("LineChart2.xlsx");
        wb.write(fileOut);
        fileOut.close();
    }
}
