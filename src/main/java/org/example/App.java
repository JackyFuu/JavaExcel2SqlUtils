package org.example;

import org.apache.poi.ss.usermodel.*;

import java.io.*;

/**
 *
 * 功能点：将excel中的数据转换为sql脚本
 * 注意点：
 * 1、excel处理：需要将excel表格设置为文本，这样输出格式就不会乱!
 * 2、通过此方法获得系统换行符：System.getProperty("line.separator", "/n");
 *
 */
public class App 
{
    public static void main( String[] args ) throws IOException {
        File targetFile = new File("G:\\BLOG\\javaUtils\\targetFile.xlsx");
        File resultFile = new File("G:\\BLOG\\javaUtils\\targetFile.sql");
        //获取系统换行符
        String lineSeparator = System.getProperty("line.separator", "/n");
        Workbook workbook = null;
        try (InputStream inputStream = new FileInputStream(targetFile); FileWriter fileWriter = new FileWriter(resultFile);) {
            //创建工作表
            workbook = WorkbookFactory.create(inputStream);
            //获取sheet对象
            Sheet sheet = workbook.getSheetAt(0);
            //总行数
            int rowLength = sheet.getLastRowNum() + 1;
            //根据第一行，获取总列数
            Row row = sheet.getRow(0);
            int colLength = row.getPhysicalNumberOfCells();
            //得到指定的单元格
            Cell SCHOOLID = null;
            Cell NAME = null;
            Cell SCHOOL_CLASSES = null;
            Cell LEGAL_TYPE = null;
            Cell CUSTOMER_NO = null;
            Cell SIGN_INSID = null;
            System.out.println("行数：" + rowLength + ",列数：" + colLength);
            for (int i = 2; i < rowLength; i++) {
                row = sheet.getRow(i);
                SCHOOLID = row.getCell(0);
                NAME = row.getCell(1);
                SIGN_INSID = row.getCell(2);
                SCHOOL_CLASSES = row.getCell(3);
                LEGAL_TYPE = row.getCell(4);
                CUSTOMER_NO = row.getCell(5);
                fileWriter.append("UPDATE zhxy_schoollist SET SCHOOL_CLASSES='" + SCHOOL_CLASSES
                        + "',LEGAL_TYPE ='"+ LEGAL_TYPE
                        + "',CUSTOMER_NO ='" + ""+ CUSTOMER_NO
                        + "',SIGN_INSID ='"+ SIGN_INSID
                        + "',NAME = trim('"+NAME+"') where trim(SCHOOLID)='"+ SCHOOLID + "';" + lineSeparator);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
