package data.utils;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.joda.time.DateTime;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * @ClassName ExcelReader
 * @Description TODO
 * @Author 多罗罗丶
 * @Date 2020/8/15 0015 16:36
 * @Version 1.0
 */
public class ExcelReader {
    private static String PATH = "E:\\桌面\\";

    /**
     * 操作03版本的xls文件：读
     */
    public static void readXls() throws Exception {
        //(1)获取文件流
        FileInputStream fis = new FileInputStream(PATH + "多罗罗的工作簿" + ".xls");
        //(2)!!!!创建工作簿:将流给到workbook对象
        Workbook workbook = new HSSFWorkbook(fis);
        //(3)获取工作表
        Sheet sheet = workbook.getSheetAt(0);
        //(4)获取行
        Row row = sheet.getRow(0);
        //(5)获取列/单元格
        Cell cell = row.getCell(0);
        //(6)尝试以字符串的形式把内容读取出来
        System.out.println(cell.getStringCellValue());
        //关流
        fis.close();
    }

    /**
     * 操作03版本的xls文件：读不同类型的数据
     * 工作中经常碰到的问题
     */
    public static void readXlsOnDifferentType() throws Exception {
        //(1)获取文件流
        FileInputStream fis = new FileInputStream(PATH + "多罗罗的工作簿" + ".xls");
        //(2)!!!!创建工作簿:将流给到workbook对象
        Workbook workbook = new HSSFWorkbook(fis);
        //(3)获取工作表
        Sheet sheet = workbook.getSheetAt(0);
        //(4)第一行通常是标题列
        Row rowTitles = sheet.getRow(0);
        //(5)智能读取表头信息
        if (rowTitles != null) {
            int cellCount = rowTitles.getPhysicalNumberOfCells();//拿到列数
            for (int cellNum = 0; cellNum < cellCount; cellNum++) {
                Cell cell = rowTitles.getCell(cellNum);
                if (cell != null) {
                    int cellType = cell.getCellType();
                    String cellValue = cell.getStringCellValue();
                    System.out.print(cellValue + " | ");
                }
            }
        }
        System.out.println();
        //(6)读取行信息
        int rowCount = sheet.getPhysicalNumberOfRows();
        for (int rowNum = 1; rowNum < rowCount; rowNum++) {
            Row rowData = sheet.getRow(rowNum);
            if (rowData != null) {
                //读取当前行每一列的数据
                int cellCount = rowTitles.getPhysicalNumberOfCells();
                for (int cellNum = 0; cellNum < cellCount; cellNum++) {
                    Cell cell = rowData.getCell(cellNum);
                    String cellValue = "";
                    //匹配列的数据类型,利用枚举类
                    if (cell != null) {
                        int cellType = cell.getCellType();
                        switch (cellType) {
                            case HSSFCell.CELL_TYPE_STRING:
                                System.out.println("这是字符串类型");
                                cellValue = cell.getStringCellValue();
                                break;
                            case HSSFCell.CELL_TYPE_BOOLEAN:
                                System.out.println("这是布尔值");
                                cellValue = String.valueOf(cell.getBooleanCellValue());
                                break;
                            case HSSFCell.CELL_TYPE_BLANK:
                                System.out.println("这是空值");
                                break;
                            case HSSFCell.CELL_TYPE_NUMERIC:
                                //数值类型可能是普通数值或者日期数值，所以要再区分
                                if (HSSFDateUtil.isCellDateFormatted(cell)) {
                                    //说明是日期
                                    System.out.println("这是数值类型之日期类型");
                                    cellValue = new DateTime(cell.getDateCellValue()).toString("yyyy-MM-dd");
                                } else {
                                    //如果是普通的数值,为了防止数字长度过长，需要转换成字符串类型的单元格才能输出
                                    System.out.println("这是数值类型");
                                    cell.setCellType(HSSFCell.CELL_TYPE_STRING);
                                    cellValue = cell.toString();
                                }
                                break;
                            case HSSFCell.CELL_TYPE_ERROR:
                                System.out.println("数据类型错误");
                                break;
                        }
                        System.out.println(cellValue);
                    }
                }
            }
        }

        //关流
        fis.close();
    }

    public static void readFormula() throws Exception {
        //(1)获取文件流
        FileInputStream fis = new FileInputStream(PATH + "多罗罗的工作簿" + ".xls");

        //(2)!!!!创建工作簿:将流给到workbook对象
        Workbook workbook = new HSSFWorkbook(fis);

        //(3)获取工作表
        Sheet sheet = workbook.getSheetAt(1);

        //(4)拿到是计算公式的单元格
        Row row = sheet.getRow(7);
        Cell cell = row.getCell(0);

        //(6)获取当前工作簿对应的计算器
        FormulaEvaluator formulaEvaluator = new HSSFFormulaEvaluator((HSSFWorkbook) workbook);

        //(5)判断是否是计算公式
        int cellType = cell.getCellType();
        switch (cellType) {
            case HSSFCell.CELL_TYPE_FORMULA:
                //如果是公式类型的单元格，我们就把公式的内容拿出来,本质是字符串
                String formula = cell.getCellFormula();
                System.out.println("公式的内容是：" + formula);
                //关键点：得知此单元格是公式之后，我们想把这个公式运算从而得到结果的话，需要获取到这个工作簿的计算器对当前单元格进行计算
                CellValue evaluate = formulaEvaluator.evaluate(cell);
                String cellValue = evaluate.formatAsString();
                System.out.println("公式的结果是：" + cellValue);
                break;
        }
        fis.close();
    }
}
