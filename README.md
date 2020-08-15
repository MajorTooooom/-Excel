# ApachePOI操作Excel

 如何操作excel的导入导出

## 导入依赖

```xml
    <dependencies>
        <!--poi依赖-start-->

        <!--xls2003版本-->
        <dependency>
            <groupId>org.apache.poi</groupId>
            <artifactId>poi</artifactId>
            <version>3.9</version>
        </dependency>

        <!--xlsx2007版本-->
        <dependency>
            <groupId>org.apache.poi</groupId>
            <artifactId>poi-ooxml</artifactId>
            <version>3.9</version>
        </dependency>

        <!--日期格式化工具-->
        <dependency>
            <groupId>joda-time</groupId>
            <artifactId>joda-time</artifactId>
            <version>2.10.1</version>
        </dependency>

        <dependency>
            <groupId>junit</groupId>
            <artifactId>junit</artifactId>
            <version>4.12</version>
        </dependency>

        <!--poi依赖-end-->
    </dependencies>
```



## 编写配置

<none/>



## 测试实例

关键接口：**WorkBook**

## WorkBook三个实现类

```bash
HSSFWorkbook：操作2003版本的xls文件
XSSFWorkbook：操作2007版本xlsx文件
	SXSSFWorkbook：属于XSSFWorkbook的增强版，也是操作2007版本的xlsx，优化了大文件的操作速度
```

2者区别：03的速度快但是行限制65535，07的慢但是不限制行。





## 基础Demo：写

只要知道excel的概念，类比就行。

关键步骤：

创建工作簿》创建工作簿》创建行》创建单元格》填充数据》IO写入

### HSSFWorkbook

```java
//一个基本的Demo
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
```



### XSSFWorkbook

和03版本的区别只是对象和文件后缀不同而已：

```java
Workbook workbook = new XSSFWorkbook();
//省略部分代码.....
FileOutputStream ops = new FileOutputStream(PATH + "多罗罗的工作簿" + ".xlsx");
```



### SXSSFWorkbook

SXSSFWorkbook的原理是默认内存100条，超出的话会先写进去，所以多了个删除临时缓存文件的步骤。和XSSFWorkbook的区别是除了实现类不同之外，还需要删除临时文件。

```java
//省略相同代码............
        //关流
        ops.close();
        //清除临时文件
        ((SXSSFWorkbook)workbook).dispose();
```



## 基础Demo：读



```java
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
```

其他实现类以此类推。



## 读取不同类型的数据

基础Demo

```java
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
```



## 公式

> 公式的思路也是先获取到cell的类型，如果是“公式”，则调用其计算方法获取结果

```java
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
```



# EasyExcel操作Excel

[官网](https://github.com/alibaba/easyexcel)也比较详细。

EasyExcel · 语雀 https://www.yuque.com/easyexcel/doc/easyexcel

设计思路个人理解：实体类进行映射、注解标示字段类型、单元格属性等；

## 实体类配置

```java
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
```



## 一行代码实现导出

```java
@Test
public void easyExcelTest01() {
    //准备好输出路径和写哪些数据
    String fileName = PATH + "使用EasyExcel生成的文件" + ".xlsx";
    List<ExamplePojo> list = EasyExcelTests.getList();

    //一行代码搞定
    EasyExcel.write(fileName, ExamplePojo.class).sheet("模板").doWrite(list);

    System.out.println("OK");
}
```



So Easy~

## 几行代码实现导入

需要配置监听器

```java
public class ExamplePojoListener extends AnalysisEventListener<ExamplePojo> {
    //.......
}
```

监听器的主要内容（1）：重写**invoke**方法

> 查出来的每条数据，如果需要自定义的操作，都写在这个方法里面，比如打印、比如持久化操作（持久化操作需要看文档详解）

```java
@Override
public void invoke(ExamplePojo data, AnalysisContext context) {
    LOGGER.info("解析到一条数据:{}", JSON.toJSONString(data));
    System.out.println(JSON.toJSONString(data));
    list.add(data);
    // 达到BATCH_COUNT了，需要去存储一次数据库，防止数据几万条数据在内存，容易OOM
    if (list.size() >= BATCH_COUNT) {
        saveData();
        // 存储完成清理 list
        list.clear();
    }
}
```

监听器的其他内容（2）：重写**doAfterAllAnalysed**方法

> 作用是读完之后要干嘛

```java
@Override
public void doAfterAllAnalysed(AnalysisContext context) {
    // 这里也要保存数据，确保最后遗留的数据也存储到数据库
    saveData();
    LOGGER.info("所有数据解析完成！");
}
```



