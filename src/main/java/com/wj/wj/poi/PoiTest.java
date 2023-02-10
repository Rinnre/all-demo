package com.wj.wj.poi;

import org.apache.poi.hpsf.DocumentSummaryInformation;
import org.apache.poi.hpsf.SummaryInformation;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;

/**
 * Created by IntelliJ IDEA
 *
 * @author wj
 * @date 2022/10/27 8:32
 */
public class PoiTest {

    public static void main(String[] args) throws IOException {
        //文件路径
        String filePath = "sample.xls";
        //创建Excel文件(Workbook)
        HSSFWorkbook workbook = new HSSFWorkbook();
        //创建工作表(Sheet)
        HSSFSheet sheet = workbook.createSheet();
        //创建工作表(Sheet)
        sheet = workbook.createSheet("Test");

        HSSFRow row = sheet.createRow(0);
        HSSFCell cell = row.createCell(0);
        cell.setCellValue("李志伟");
        row.createCell(1).setCellValue(false);
        row.createCell(2).setCellValue(new Date());
        row.createCell(3).setCellValue(12.345);

        //创建文档信息
        workbook.createInformationProperties();
        //摘要信息
        DocumentSummaryInformation dsi = workbook.getDocumentSummaryInformation();
        dsi.setCategory("类别:Excel文件");
        dsi.setManager("管理者:李志伟");
        dsi.setCompany("公司:--");
        //摘要信息
        SummaryInformation si = workbook.getSummaryInformation();
        si.setSubject("主题:--");
        si.setTitle("标题:测试文档");
        si.setAuthor("作者:李志伟");
        si.setComments("备注:POI测试文档");


        FileOutputStream out = new FileOutputStream(filePath);
        //保存Excel文件
        workbook.write(out);
        out.close();//关闭文件流
        System.out.println("OK!");
    }

}
