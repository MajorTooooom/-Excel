import com.alibaba.excel.EasyExcel;
import listener.ExamplePojoListener;
import org.junit.Test;
import pojo.ExamplePojo;

import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * @ClassName EasyExcelTests
 * @Description TODO 使用EasyExcel操作表格读写
 * @Author 多罗罗丶
 * @Date 2020/8/15 0015 19:19
 * @Version 1.0
 * 主要参考：EasyExcel · 语雀 https://www.yuque.com/easyexcel/doc/easyexcel
 * 官网文档： https://github.com/alibaba/easyexcel
 */
public class EasyExcelTests {
    private static String PATH = "E:\\桌面\\";

    /**
     * 测试EasyExcel
     */
    @Test
    public void easyExcelWriteTest01() {
        //准备好输出路径和写哪些数据
        String fileName = PATH + "使用EasyExcel生成的文件" + ".xlsx";
        List<ExamplePojo> list = EasyExcelTests.getList();

        //一行代码搞定
        EasyExcel.write(fileName, ExamplePojo.class).sheet("模板").doWrite(list);

        System.out.println("OK");
    }


    /**
     * 详细的读取excel的操作文档：https://www.yuque.com/easyexcel/doc/read
     */
    @Test
    public void easyExcelReadTest01() {
        String fileName = PATH + "使用EasyExcel生成的文件" + ".xlsx";
        EasyExcel.read(fileName, ExamplePojo.class, new ExamplePojoListener()).sheet().doRead();
    }


    /**
     * 准备一些测试数据
     */
    public static List<ExamplePojo> getList() {
        List<ExamplePojo> list = new ArrayList<ExamplePojo>();
        ExamplePojo temp;
        for (int i = 0; i < 100; i++) {
            temp = new ExamplePojo();
            temp.setBn(i <= 50 ? true : false);
            temp.setDate(new Date());
            temp.setName("名字" + i);
            temp.setNumber(BigDecimal.valueOf(i));
            list.add(temp);
        }
        return list;
    }


}
