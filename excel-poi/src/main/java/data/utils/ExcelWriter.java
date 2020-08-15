package data.utils;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import java.io.FileOutputStream;

/**
 * @ClassName ExcelWriter
 * @Description TODO 使用ApachePOI对Excel进行读写操作
 * @Author 多罗罗丶
 * @Date 2020/8/15 0015 15:48
 * @Version 1.0
 */
public class ExcelWriter {
    private static String PATH = "E:\\桌面\\";


    /**
     * 操作03版本的xls文件：写
     */
    public static void writeXls() throws Exception {
        //(一)创建工作簿
        Workbook workbook = new HSSFWorkbook();
        //(二)创建工作表
        Sheet sheet = workbook.createSheet("多罗罗的工作表");
        //(三)创建工作行
        Row row1 = sheet.createRow(0);
        //(四)创建单元格
        Cell cell1 = row1.createCell(0);
        //(五)填充数据
        cell1.setCellValue("这是第一行第一格");

        //测试多一行
        Row row2 = sheet.createRow(1);
        Cell cell2 = row2.createCell(0);
        cell2.setCellValue(new DateTime().toString("yyyy-MM-dd HH:mm:ss"));

        //(六)使用IO流生成表，03版本的文件是xls结尾
        FileOutputStream ops = new FileOutputStream(PATH + "多罗罗的工作簿" + ".xls");
        //工作簿对象使用流将工作簿写出
        workbook.write(ops);
        //关流
        ops.close();
    }

    /**
     * 操作07版本的xlsx文件：写
     */
    public static void writeXlsx() throws Exception {
        //(一)创建工作簿
        Workbook workbook = new XSSFWorkbook();
        //(二)创建工作表
        Sheet sheet = workbook.createSheet("多罗罗的工作表");
        //(三)创建工作行
        Row row1 = sheet.createRow(0);
        //(四)创建单元格
        Cell cell1 = row1.createCell(0);
        //(五)填充数据
        cell1.setCellValue("这是第一行第一格");

        //测试多一行
        Row row2 = sheet.createRow(1);
        Cell cell2 = row2.createCell(0);
        cell2.setCellValue(new DateTime().toString("yyyy-MM-dd HH:mm:ss"));

        //(六)使用IO流生成表，03版本的文件是xls结尾
        FileOutputStream ops = new FileOutputStream(PATH + "多罗罗的工作簿" + ".xlsx");
        //工作簿对象使用流将工作簿写出
        workbook.write(ops);
        //关流
        ops.close();
    }

    /**
     * 操作07版本的xlsx文件：写
     * 区别是删除临时文件
     */
    public static void writeXlsxS() throws Exception {
        //(一)创建工作簿
        Workbook workbook = new SXSSFWorkbook();
        //(二)创建工作表
        Sheet sheet = workbook.createSheet("多罗罗的工作表");
        //(三)创建工作行
        Row row1 = sheet.createRow(0);
        //(四)创建单元格
        Cell cell1 = row1.createCell(0);
        //(五)填充数据
        cell1.setCellValue("这是第一行第一格");

        //测试多一行
        Row row2 = sheet.createRow(1);
        Cell cell2 = row2.createCell(0);
        cell2.setCellValue(new DateTime().toString("yyyy-MM-dd HH:mm:ss"));

        //(六)使用IO流生成表，03版本的文件是xls结尾
        FileOutputStream ops = new FileOutputStream(PATH + "多罗罗的工作簿S" + ".xlsx");
        //工作簿对象使用流将工作簿写出
        workbook.write(ops);
        //关流
        ops.close();
        //清除临时文件
        ((SXSSFWorkbook) workbook).dispose();
    }
}
