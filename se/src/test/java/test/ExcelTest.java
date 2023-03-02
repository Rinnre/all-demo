package test;

import com.wj.poi.entity.ExcelChartTypeEnum;
import com.wj.poi.entity.ExcelEntity;
import com.wj.poi.utils.ExcelUtil3;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

/**
 * Created by IntelliJ IDEA
 *
 * @author wj
 * @date 2022/12/10 11:58
 */

public class ExcelTest {

    public static void main(String[] args) throws IOException {
        List<ExcelEntity> excelEntities = new ArrayList<>();
        ExcelEntity excelEntity = new ExcelEntity();

        excelEntity.setTableName("智能电表1类比智能电表2");


        LinkedHashMap<String, String> tableFiled = new LinkedHashMap<>();
        tableFiled.put("time", "时间");
        tableFiled.put("electricity_01", "电表1");
        tableFiled.put("electricity_02", "电表2");
        excelEntity.setTableFiled(tableFiled);

        Map<String, List<Object>> tableDataList = new HashMap<>();

        Set<String> filedSet = tableFiled.keySet();
        for (String filed : filedSet) {
            List<Object> list = new ArrayList<>();
            for (int i = 1; i < 24; i++) {
                if ("time".equals(filed)) {
                    String time = "2022-10-28 " + i + ":00:00";
                    list.add(time);
                } else {
                    list.add(new Random().nextInt(3));

                }
            }
            tableDataList.put(filed, list);

        }


        excelEntity.setExcelChartTypeEnum(ExcelChartTypeEnum.PIE);
        excelEntity.setTableHeadTime("2022-10");

        Map<String, Object> filedUnit = new HashMap<>();
        filedUnit.put("x", "时间");
        filedUnit.put("y", "电能/k");
        excelEntity.setFiledUnit(filedUnit);
        excelEntity.setStatistics(true);
        excelEntity.setTableDataList(tableDataList);


        excelEntities.add(excelEntity);


        excelEntity = new ExcelEntity();

        excelEntity.setTableName("智能气表1类比智能气表2");
        excelEntity.setTableHeadTime("2022-10");


        tableFiled = new LinkedHashMap<>();
        tableFiled.put("time", "时间");
        tableFiled.put("gas_01", "智能燃气表1");
        tableFiled.put("gas_02", "智能燃气表2");
        excelEntity.setTableFiled(tableFiled);

        tableDataList = new HashMap<>();

        filedSet = tableFiled.keySet();
        for (String filed : filedSet) {
            List<Object> list = new ArrayList<>();
            for (int i = 1; i < 31; i++) {
                if ("time".equals(filed)) {
                    String time = "2022-10-" + i;
                    list.add(time);
                } else {
                    list.add(new java.util.Random().nextDouble() * 10d);

                }
            }
            tableDataList.put(filed, list);

        }
        excelEntity.setTableDataList(tableDataList);
        excelEntity.setExcelChartTypeEnum(ExcelChartTypeEnum.BAR_CHAR);
        filedUnit = new HashMap<>();
        filedUnit.put("x", "时间");
        filedUnit.put("y", "气能/(m^3)");
        excelEntity.setFiledUnit(filedUnit);
        excelEntity.setStatistics(true);
        excelEntities.add(excelEntity);
        excelEntity = new ExcelEntity();

        excelEntity.setTableName("智能水表1类比智能水表2");


        tableFiled = new LinkedHashMap<>();
        tableFiled.put("time", "时间");
        tableFiled.put("water_01", "智能水表1");
        tableFiled.put("water_02", "智能水表2");
        excelEntity.setTableFiled(tableFiled);

        tableDataList = new HashMap<>();

        filedSet = tableFiled.keySet();
        for (String filed : filedSet) {
            List<Object> list = new ArrayList<>();
            for (int i = 1; i < 13; i++) {
                if ("time".equals(filed)) {
                    String time = "2022-" + i;
                    list.add(time);
                } else {
                    list.add(new java.util.Random().nextDouble() * 10d);

                }
            }
            tableDataList.put(filed, list);

        }
        excelEntity.setTableDataList(tableDataList);
        excelEntity.setTableHeadTime("2022-10");
        excelEntity.setExcelChartTypeEnum(ExcelChartTypeEnum.PIE);
        filedUnit = new HashMap<>();
        filedUnit.put("x", "时间");
        filedUnit.put("y", "水能/(m^3)");
        excelEntity.setStatistics(true);
        excelEntities.add(excelEntity);


        XSSFWorkbook excelTable = new XSSFWorkbook();
        ExcelUtil3.getExcelTable(excelTable, excelEntities);


        String filename = "历史数据.xlsx";
        FileOutputStream outputStream = new FileOutputStream(filename);
        excelTable.write(outputStream);
        outputStream.close();
        System.out.println("打印完成");
    }

}
