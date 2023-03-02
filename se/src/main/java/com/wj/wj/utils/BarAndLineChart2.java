package com.wj.wj.utils;


import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Chart;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.*;
import org.openxmlformats.schemas.drawingml.x2006.chart.*;

import java.io.FileOutputStream;

public class BarAndLineChart2 {

    public static void main(String[] args) throws Exception {
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet("Sheet1");

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
            cell.setCellValue(new java.util.Random().nextDouble()*10d);
        }

        XSSFDrawing drawing = sheet.createDrawingPatriarch();
        XSSFClientAnchor anchor = (XSSFClientAnchor) drawing.createAnchor(0, 0, 0, 0, 4, 0, 11, 15);

        Chart chart = drawing.createChart(anchor);

        CTChart ctChart = ((XSSFChart)chart).getCTChart();
        CTPlotArea ctPlotArea = ctChart.getPlotArea();

        //the bar chart
        CTBarChart ctBarChart = ctPlotArea.addNewBarChart();
        CTBoolean ctBoolean = ctBarChart.addNewVaryColors();
        ctBoolean.setVal(false);
        ctBarChart.addNewBarDir().setVal(STBarDir.COL);

        //the bar series
        CTBarSer ctBarSer = ctBarChart.addNewSer();
        CTSerTx ctSerTx = ctBarSer.addNewTx();
        CTStrRef ctStrRef = ctSerTx.addNewStrRef();
        ctBarSer.addNewIdx().setVal(0);
        CTAxDataSource cttAxDataSource = ctBarSer.addNewCat();
        ctStrRef = cttAxDataSource.addNewStrRef();
        CTNumDataSource ctNumDataSource = ctBarSer.addNewVal();
        CTNumRef ctNumRef = ctNumDataSource.addNewNumRef();

        //at least the border lines in Libreoffice Calc ;-)
        ctBarSer.addNewSpPr().addNewLn().addNewSolidFill().addNewSrgbClr().setVal(new byte[] {0,0,0});

        //telling the BarChart that it has axes and giving them Ids
        ctBarChart.addNewAxId().setVal(123456); //cat axis 1 (bars)
        ctBarChart.addNewAxId().setVal(123457); //val axis 1 (left)

        //the line chart
        CTLineChart ctLineChart = ctPlotArea.addNewLineChart();
        ctBoolean = ctLineChart.addNewVaryColors();
        ctBoolean.setVal(true);

        //the line series
        CTLineSer ctLineSer = ctLineChart.addNewSer();
        ctSerTx = ctLineSer.addNewTx();
        ctStrRef = ctSerTx.addNewStrRef();
        ctStrRef.setF("Sheet1!$B$1");
        ctLineSer.addNewIdx().setVal(1);
        cttAxDataSource = ctLineSer.addNewCat();
        ctStrRef = cttAxDataSource.addNewStrRef();
        ctStrRef.setF("Sheet1!$A$2:$A$7");
        ctNumDataSource = ctLineSer.addNewVal();
        ctNumRef = ctNumDataSource.addNewNumRef();
        ctNumRef.setF("Sheet1!$B$2:$C$7");

        //at least the border lines in Libreoffice Calc ;-)
        ctLineSer.addNewSpPr().addNewLn().addNewSolidFill().addNewSrgbClr().setVal(new byte[] {0,0,0});

        //telling the LineChart that it has axes and giving them Ids
        ctLineChart.addNewAxId().setVal(123458); //cat axis 2 (lines)
        ctLineChart.addNewAxId().setVal(123459); //val axis 2 (right)

        //cat axis 1 (bars)
        CTCatAx ctCatAx = ctPlotArea.addNewCatAx();
        ctCatAx.addNewAxId().setVal(123456); //id of the cat axis
        CTScaling ctScaling = ctCatAx.addNewScaling();
        ctCatAx.addNewDelete().setVal(false);
        ctCatAx.addNewAxPos().setVal(STAxPos.B);
        ctCatAx.addNewCrossAx().setVal(123457); //id of the val axis

        //val axis 1 (left)
        CTValAx ctValAx = ctPlotArea.addNewValAx();
        ctValAx.addNewAxId().setVal(123457); //id of the val axis
        ctScaling = ctValAx.addNewScaling();
        ctScaling.addNewOrientation().setVal(STOrientation.MIN_MAX);
        ctValAx.addNewDelete().setVal(false);
        ctValAx.addNewAxPos().setVal(STAxPos.L);
        ctValAx.addNewCrossAx().setVal(123456); //id of the cat axis
        ctValAx.addNewCrosses().setVal(STCrosses.AUTO_ZERO); //this val axis crosses the cat axis at zero
        ctValAx.addNewTickLblPos().setVal(STTickLblPos.NEXT_TO);

        //cat axis 2 (lines)
        ctCatAx = ctPlotArea.addNewCatAx();
        ctCatAx.addNewAxId().setVal(123458); //id of the cat axis
        ctScaling = ctCatAx.addNewScaling();
        ctScaling.addNewOrientation().setVal(STOrientation.MIN_MAX);
        ctCatAx.addNewDelete().setVal(true); //this cat axis is deleted
        ctCatAx.addNewAxPos().setVal(STAxPos.B);
        ctCatAx.addNewCrossAx().setVal(123459); //id of the val axis
        ctCatAx.addNewTickLblPos().setVal(STTickLblPos.NEXT_TO);

        //val axis 2 (right)
        ctValAx = ctPlotArea.addNewValAx();
        ctValAx.addNewAxId().setVal(123459); //id of the val axis
        ctScaling = ctValAx.addNewScaling();
        ctValAx.addNewAxPos().setVal(STAxPos.R);
        ctValAx.addNewCrossAx().setVal(123458); //id of the cat axis

        //legend
        CTLegend ctLegend = ctChart.addNewLegend();
        ctLegend.addNewLegendPos().setVal(STLegendPos.B);
        ctLegend.addNewOverlay().setVal(false);

        FileOutputStream fileOut = new FileOutputStream("BarAndLineChart2.xlsx");
        wb.write(fileOut);
        fileOut.close();
    }
}
