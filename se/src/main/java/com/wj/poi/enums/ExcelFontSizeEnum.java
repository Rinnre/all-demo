package com.wj.poi.enums;

/**
 * Created by IntelliJ IDEA
 *
 * @author wj
 * @date 2022/12/21 17:21
 */
public enum ExcelFontSizeEnum {
    DEFAULT((short)15);
    private Short fontSize;

    ExcelFontSizeEnum(Short fontSize) {
        this.fontSize = fontSize;
    }

    public Short getFontSize() {
        return fontSize;
    }
}
