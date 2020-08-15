import data.utils.ExcelReader;
import data.utils.ExcelWriter;
import org.junit.Test;

/**
 * @ClassName ExcelWriteTests
 * @Description TODO 使用ApachePOI对Excel进行读写操作
 * @Author 多罗罗丶
 * @Date 2020/8/15 0015 15:42
 * @Version 1.0
 */
public class ExcelWriteTests {

    @Test
    public void writeTests() throws Exception {
        ExcelWriter.writeXls();//03
        ExcelWriter.writeXlsx();//07
    }

    @Test
    public void readTests() throws Exception {
        ExcelReader.readXls();
    }

    @Test
    public void readDifferentTypeTests() throws Exception {
        ExcelReader.readXlsOnDifferentType();
    }

    @Test
    public void readFormulaTests() throws Exception {
        ExcelReader.readFormula();
    }

}
