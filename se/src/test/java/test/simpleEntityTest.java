package test;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

/**
 * Created by IntelliJ IDEA
 *
 * @author wj
 * @date 2022/12/10 15:42
 */
public class simpleEntityTest {
    public static void main(String[] args) throws IOException {
        //记录导出65536行数据多长时间
        long begin = System.currentTimeMillis();
        //创建一个工作簿
        XSSFWorkbook workbook = new XSSFWorkbook();
        //创建表
        XSSFSheet sheet = workbook.createSheet();
        //写入数据  此时可以写入的数据可以超过65536
        // 时间行为表头单独抽出(以上可以抽出为导出模板)
        XSSFRow firstRow = sheet.createRow(0);
        for (int col = 1; col < 32; col++) {
            XSSFCell cell = firstRow.createCell(col);
            cell.setCellValue(col - 1 + "天");
        }

//        行数
        for (int rowNumber = 1; rowNumber < 73; rowNumber++) {
            XSSFRow row = sheet.createRow(rowNumber);
//            列数
            for (int cellNum = 0; cellNum < 32; cellNum++) {
                // 第一列为设备名称+id
                XSSFCell cell = row.createCell(cellNum);
                if (cellNum == 0) {
                    cell.setCellValue("设备名称" + cellNum);
                } else {
                    cell.setCellValue(cellNum*1000+".00");
                }
                sheet.autoSizeColumn(cellNum,true);
            }
        }
        FileOutputStream outputStream = new FileOutputStream("测试07.xlsx");
        workbook.write(outputStream);
        //关闭资源
        outputStream.close();
        long end = System.currentTimeMillis();
    }
}
