package com.wj.poi.entity;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;

/**
 * Created by IntelliJ IDEA
 *
 * @author wj
 * @date 2022/12/21 18:15
 */
public class ExcelTableStyleEntity {



    private boolean isWrapText;

    /**
     *  水平对齐
     *  GENERAL,
     *  LEFT,
     *  CENTER,
     *  RIGHT,
     *  FILL,
     *  JUSTIFY,
     *  CENTER_SELECTION,
     *  DISTRIBUTED;
     */
    private HorizontalAlignment alignment;

    /**
     * 垂直对齐
     */
    private VerticalAlignment verticalAlignment;

    private Short fillForegroundColor;

    private Short fillBackgroundColor;

    private FillPatternType fillPatternType;

    private ExcelFontStyleEntity excelFontStyleEntity;

    /**
     * 边框设置
     *
     */
    private BorderStyle borderStyle;

    private Short borderColor;

    private CellRangeAddress cellAddresses;

    private XSSFSheet sheet;

    public BorderStyle getBorderStyle() {
        return borderStyle;
    }

    public void setBorderStyle(BorderStyle borderStyle) {
        this.borderStyle = borderStyle;
    }

    public Short getBorderColor() {
        return borderColor;
    }

    public void setBorderColor(Short borderColor) {
        this.borderColor = borderColor;
    }

    public CellRangeAddress getCellAddresses() {
        return cellAddresses;
    }

    public void setCellAddresses(CellRangeAddress cellAddresses) {
        this.cellAddresses = cellAddresses;
    }

    public XSSFSheet getSheet() {
        return sheet;
    }

    public void setSheet(XSSFSheet sheet) {
        this.sheet = sheet;
    }

    public ExcelFontStyleEntity getExcelFontStyleEntity() {
        return excelFontStyleEntity;
    }

    public void setExcelFontStyleEntity(ExcelFontStyleEntity excelFontStyleEntity) {
        this.excelFontStyleEntity = excelFontStyleEntity;
    }

    public Short getFillForegroundColor() {
        return fillForegroundColor;
    }

    public void setFillForegroundColor(Short fillForegroundColor) {
        this.fillForegroundColor = fillForegroundColor;
    }

    public Short getFillBackgroundColor() {
        return fillBackgroundColor;
    }

    public void setFillBackgroundColor(Short fillBackgroundColor) {
        this.fillBackgroundColor = fillBackgroundColor;
    }

    public FillPatternType getFillPatternType() {
        return fillPatternType;
    }

    public void setFillPatternType(FillPatternType fillPatternType) {
        this.fillPatternType = fillPatternType;
    }



    public boolean isWrapText() {
        return isWrapText;
    }

    public void setWrapText(boolean wrapText) {
        isWrapText = wrapText;
    }

    public HorizontalAlignment getAlignment() {
        return alignment;
    }

    public void setAlignment(HorizontalAlignment alignment) {
        this.alignment = alignment;
    }

    public VerticalAlignment getVerticalAlignment() {
        return verticalAlignment;
    }

    public void setVerticalAlignment(VerticalAlignment verticalAlignment) {
        this.verticalAlignment = verticalAlignment;
    }
}
