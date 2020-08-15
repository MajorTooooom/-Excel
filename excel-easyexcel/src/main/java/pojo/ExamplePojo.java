package pojo;

import com.alibaba.excel.annotation.ExcelIgnore;
import com.alibaba.excel.annotation.ExcelProperty;
import lombok.Data;

import java.math.BigDecimal;
import java.util.Date;

/**
 * @ClassName ExamplePojo
 * @Description TODO EasyExcel使用实体类进行映射
 * @Author 多罗罗丶
 * @Date 2020/8/15 0015 19:03
 * @Version 1.0
 */
@Data
public class ExamplePojo {
    /**
     * 设置EasyExcel进行excel导入导出时，是否对当前属性进行处理（是否是表格列）
     * value: 标题行的内容
     * index: 下标（顺序）
     */
    @ExcelProperty(value = "日期", index = 0)
    private Date date;

    /**
     * 名称
     */
    @ExcelProperty(value = "名称", index = 1)
    private String name;

    /**
     * 数值类型
     */
    @ExcelProperty(value = "数值", index = 2)
    private BigDecimal number;

    /**
     * 布尔类型
     */
    @ExcelProperty(value = "布尔类型", index = 3)
    private Boolean bn;

    /**
     * //TODO 设置当前实体类属性非excel的字段，那么在导入导出的时候会被忽略掉
     */
    @ExcelIgnore
    private String notForExcel;
}
