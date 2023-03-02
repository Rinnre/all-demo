package com.wj.poi.entity;

import org.apache.poi.ss.usermodel.FontFormatting;

/**
 * Created by IntelliJ IDEA
 *
 * @author wj
 * @date 2022/12/21 16:51
 */
public class ExcelFontStyleEntity {
    private String fontFamily;

    private Short fontSize;

    private Short fontColor;

    private Byte fontUnderLine;

    private Byte fontTypeOffset;

    private boolean fontStrikeout;

    private boolean isFontBold;

    public Boolean getFontBold() {
        return isFontBold;
    }

    public void setFontBold(Boolean fontBold) {
        isFontBold = fontBold;
    }

    public String getFontFamily() {
        return fontFamily;
    }

    public void setFontFamily(String fontFamily) {
        this.fontFamily = fontFamily;
    }

    public Short getFontSize() {
        return fontSize;
    }

    public void setFontSize(Short fontSize) {
        this.fontSize = fontSize;
    }

    public Short getFontColor() {
        return fontColor;
    }

    public void setFontColor(Short fontColor) {
        this.fontColor = fontColor;
    }

    public Byte getFontUnderLine() {
        return fontUnderLine;
    }

    public void setFontUnderLine(Byte fontUnderLine) {
        this.fontUnderLine = fontUnderLine;
    }

    public Byte getFontTypeOffset() {
        return fontTypeOffset;
    }

    public void setFontTypeOffset(Byte fontTypeOffset) {
        this.fontTypeOffset = fontTypeOffset;
    }

    public Boolean getFontStrikeout() {
        return fontStrikeout;
    }

    public void setFontStrikeout(Boolean fontStrikeout) {
        this.fontStrikeout = fontStrikeout;
    }
}
