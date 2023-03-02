package com.wj.poi.enums;

/**
 * Created by IntelliJ IDEA
 *
 * @author wj
 * @date 2022/12/21 17:22
 */
public enum ExcelFontFamilyEnum {
    DEFAULT("宋体");
    private String fontFamily;

    ExcelFontFamilyEnum(String fontFamily) {
        this.fontFamily = fontFamily;
    }

    public String getFontFamily() {
        return fontFamily;
    }
}
