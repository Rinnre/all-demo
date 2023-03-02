package com.wj.excel.util;

/**
 * Created by IntelliJ IDEA
 *
 * @author wj
 * @date 2022/10/31 8:22
 */
public enum ExcelChartTypeEnum {
    LINE_CHAR("折线图",1),
    PIE("饼图",2),
    BAR_CHAR("柱状图",3);

    private String name;
    private Integer type;

    private ExcelChartTypeEnum(String name, Integer type){
        this.name = name;
        this.type = type;
    }


    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public void setType(Integer type) {
        this.type = type;
    }

    public Integer getType() {
        return type;
    }
}
