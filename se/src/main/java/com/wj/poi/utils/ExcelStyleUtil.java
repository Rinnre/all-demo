package com.wj.poi.utils;

import com.wj.poi.entity.ExcelFontStyleEntity;
import com.wj.poi.entity.ExcelTableStyleEntity;
import com.wj.poi.enums.ExcelFontFamilyEnum;
import com.wj.poi.enums.ExcelFontSizeEnum;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;


/**
 * Created by IntelliJ IDEA
 *
 * @author wj
 * @date 2022/12/21 18:10
 */
public class ExcelStyleUtil {

    /**
     * 生成默认单元格样式表头
     * 表头默认样式
     *
     * @param workbook 工作簿
     * @return {@link CellStyle}
     */
    public static CellStyle generateDefaultTableHeadCellStyle(Workbook workbook,XSSFSheet sheet,CellRangeAddress cellAddresses) {
        ExcelTableStyleEntity excelTableHeadStyleEntity = new ExcelTableStyleEntity();
        ExcelFontStyleEntity excelTableHeadFontStyleEntity = new ExcelFontStyleEntity();

        // 字体设置
        excelTableHeadFontStyleEntity.setFontBold(true);
        excelTableHeadFontStyleEntity.setFontFamily(excelTableHeadFontStyleEntity.getFontFamily());
        excelTableHeadFontStyleEntity.setFontColor(IndexedColors.WHITE.getIndex());
        excelTableHeadFontStyleEntity.setFontSize(excelTableHeadFontStyleEntity.getFontSize());

        // 填充背景
        excelTableHeadStyleEntity.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
        excelTableHeadStyleEntity.setFillPatternType(FillPatternType.SOLID_FOREGROUND);
        // 自动换行
        excelTableHeadStyleEntity.setWrapText(true);

        excelTableHeadStyleEntity.setExcelFontStyleEntity(excelTableHeadFontStyleEntity);

        // 设置边框
        excelTableHeadStyleEntity.setBorderStyle(BorderStyle.THIN);
        excelTableHeadStyleEntity.setBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        excelTableHeadStyleEntity.setCellAddresses(cellAddresses);
        excelTableHeadStyleEntity.setSheet(sheet);

        return generateCellStyle(workbook, excelTableHeadStyleEntity);
    }




    /**
     * 自定义单元格样式
     *
     * @return {@link XSSFCellStyle}
     */
    public static CellStyle generateCellStyle(Workbook workbook, ExcelTableStyleEntity excelTableStyleEntity) {
        CellStyle cellStyle = workbook.createCellStyle();


        // 设置文字对齐
        if (null != excelTableStyleEntity.getAlignment()) {
            excelTableStyleEntity.setAlignment(excelTableStyleEntity.getAlignment());
        } else {
            // 默认居中
            excelTableStyleEntity.setAlignment(HorizontalAlignment.CENTER);
        }
        if (null != excelTableStyleEntity.getVerticalAlignment()) {
            excelTableStyleEntity.setVerticalAlignment(excelTableStyleEntity.getVerticalAlignment());
        } else {
            excelTableStyleEntity.setVerticalAlignment(VerticalAlignment.CENTER);

        }

        // 设置填充颜色
        if (null != excelTableStyleEntity.getFillPatternType()) {
            cellStyle.setFillPattern(excelTableStyleEntity.getFillPatternType());
        }

        if (null != excelTableStyleEntity.getFillForegroundColor()) {
            cellStyle.setFillForegroundColor(excelTableStyleEntity.getFillForegroundColor());
        }

        if (null != excelTableStyleEntity.getFillBackgroundColor()) {
            cellStyle.setFillBackgroundColor(excelTableStyleEntity.getFillBackgroundColor());
        }

        // 是否自动换行
        cellStyle.setWrapText(excelTableStyleEntity.isWrapText());

        // 设置字体颜色
        Font font = generateFontStyle(workbook, excelTableStyleEntity.getExcelFontStyleEntity());
        cellStyle.setFont(font);

        // 边框设置
        setTableBorder(cellStyle,excelTableStyleEntity.getBorderStyle(),
                excelTableStyleEntity.getBorderColor(),excelTableStyleEntity.getCellAddresses(), excelTableStyleEntity.getSheet());

        return cellStyle;
    }


    /**
     * 生成字体样式
     *
     * @param workbook             工作簿
     * @param excelFontStyleEntity excel字体样式实体
     * @return {@link Font}
     */
    private static Font generateFontStyle(Workbook workbook, ExcelFontStyleEntity excelFontStyleEntity) {
        // 创建字体对象
        Font font = workbook.createFont();

        // todo 重构为反射填充（解决方法名称映射问题）

        // 设置字体名称
        if (null != excelFontStyleEntity.getFontFamily()) {
            font.setFontName(excelFontStyleEntity.getFontFamily());
        } else {
            font.setFontName(ExcelFontFamilyEnum.DEFAULT.getFontFamily());
        }

        // 设置字号
        if (null != excelFontStyleEntity.getFontSize()) {
            font.setFontHeightInPoints(excelFontStyleEntity.getFontSize());
        } else {
            font.setFontHeightInPoints(ExcelFontSizeEnum.DEFAULT.getFontSize());
        }

        // 设置字体颜色
        if (null != excelFontStyleEntity.getFontColor()) {
            font.setColor(excelFontStyleEntity.getFontColor());
        } else {
            font.setColor(HSSFFont.COLOR_NORMAL);
        }

        // 字体是否加粗
        if (null != excelFontStyleEntity.getFontBold()) {
            font.setBold(excelFontStyleEntity.getFontBold());
        }

        //设置下划线
        if (null != excelFontStyleEntity.getFontUnderLine()) {
            font.setUnderline(excelFontStyleEntity.getFontUnderLine());
        }

        //设置上标下标
        if (null != excelFontStyleEntity.getFontTypeOffset()) {
            font.setTypeOffset(excelFontStyleEntity.getFontTypeOffset());
        }

        //设置删除线
        if (null != excelFontStyleEntity.getFontStrikeout()) {
            font.setStrikeout(excelFontStyleEntity.getFontStrikeout());
        }
        return font;
    }


    /**
     * 边界设置表
     *
     * @param cellStyle     单元格样式
     * @param borderStyle   边框样式
     * @param borderColor   边框颜色
     * @param cellAddresses 单元格地址
     * @param sheet         表
     */
    private static void setTableBorder(CellStyle cellStyle, BorderStyle borderStyle, short borderColor,
                                       CellRangeAddress cellAddresses, XSSFSheet sheet) {

        // 普通单元格设置边框以及边框颜色
        if (null != cellStyle) {
            cellStyle.setBorderBottom(borderStyle);
            cellStyle.setBorderLeft(borderStyle);
            cellStyle.setBorderRight(borderStyle);
            cellStyle.setBorderTop(borderStyle);
            cellStyle.setBottomBorderColor(borderColor);
            cellStyle.setLeftBorderColor(borderColor);
            cellStyle.setRightBorderColor(borderColor);
            cellStyle.setTopBorderColor(borderColor);
        }

        // 合并单元格设置边框以及边框颜色
        if (null != cellAddresses && null != sheet) {
            RegionUtil.setBorderTop(borderStyle, cellAddresses, sheet);
            RegionUtil.setBorderRight(borderStyle, cellAddresses, sheet);
            RegionUtil.setBorderLeft(borderStyle, cellAddresses, sheet);
            RegionUtil.setBorderBottom(borderStyle, cellAddresses, sheet);
            RegionUtil.setBottomBorderColor(borderColor, cellAddresses, sheet);
            RegionUtil.setLeftBorderColor(borderColor, cellAddresses, sheet);
            RegionUtil.setRightBorderColor(borderColor, cellAddresses, sheet);
            RegionUtil.setTopBorderColor(borderColor, cellAddresses, sheet);
        }
    }
}
