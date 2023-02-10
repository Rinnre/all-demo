package com.wj.excel.util;

import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * Created by IntelliJ IDEA
 *
 * @author wj
 * @date 2022/10/28 13:44
 */
public class ExcelEntity {

    /**
     *  sheet以及表名
     */
    private String tableName;

    /**
     *
     * [
     *  {
     *      "time":"时间",
     *      "gas_01":"表1"
     *  }
     * ]
     *
     */
    private LinkedHashMap<String, String> tableFiled;

    /**
     * 饼图不需要设置
     * 设置x轴和y轴的标题
     * ["x":"时间","y":"水能/(m^3")]
     */
    private Map<String, Object> FiledUnit;

    /**
     * 按时间排好顺序的值
     * 当前时间无数据的需要补0
     * <p>
     * {
     * "time":[],
     * "table1":[],
     * "table2":[]
     * }
     */
    private Map<String, List<Object>> tableDataList;

    /**
     * 不需要生成图则设为null
     * 类型：折线图，柱状图，饼图
     */

    private ExcelChartTypeEnum excelChartTypeEnum;

    /**
     * 是否需要统计总用量
     */
    private boolean isStatistics;

    public boolean getStatistics() {
        return isStatistics;
    }

    public void setStatistics(boolean statistics) {
        isStatistics = statistics;
    }

    public String getTableName() {
        return tableName;
    }

    public void setTableName(String tableName) {
        this.tableName = tableName;
    }


    public LinkedHashMap<String, String> getTableFiled() {
        return tableFiled;
    }

    public void setTableFiled(LinkedHashMap<String, String> tableFiled) {
        this.tableFiled = tableFiled;
    }

    public Map<String, List<Object>> getTableDataList() {
        return tableDataList;
    }

    public void setTableDataList(Map<String, List<Object>> tableDataList) {
        this.tableDataList = tableDataList;
    }

    public Map<String, Object> getFiledUnit() {
        return FiledUnit;
    }

    public void setFiledUnit(Map<String, Object> filedUnit) {
        FiledUnit = filedUnit;
    }

    public ExcelChartTypeEnum getExcelChartTypeEnum() {
        return excelChartTypeEnum;
    }

    public void setExcelChartTypeEnum(ExcelChartTypeEnum excelChartTypeEnum) {
        this.excelChartTypeEnum = excelChartTypeEnum;
    }
}
