package com.wj.poi.utils;

import org.apache.poi.hssf.usermodel.HSSFFooter;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFPrintSetup;
import org.apache.poi.xssf.usermodel.XSSFSheet;

/**
 * Created by IntelliJ IDEA
 *
 * @author wj
 * @date 2022/12/21 16:42
 */
public class ExcelUtil {

    /**
     * 数据填充（实体类反射、自定义构造）
     */
    public static void fillDataIntoTable(){


    }


    /**
     * 配置默认打印设置
     *
     * @param sheet 表
     */
    private static void defaultPrintSetting(XSSFSheet sheet) {
        XSSFPrintSetup printSetup = sheet.getPrintSetup();
        // 横纵向设置（纵向）
        printSetup.setLandscape(false);

        // 设置纸张（A3）
        printSetup.setPaperSize(PrintSetup.A3_PAPERSIZE);

        // 设置将工作表调整为一页显示
        sheet.setFitToPage(true);

        // 页码设置
        Footer footer = sheet.getFooter();
        footer.setCenter("第" + HSSFFooter.page() + "页，共" + HSSFFooter.numPages() + "页");

    }

}
