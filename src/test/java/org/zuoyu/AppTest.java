package org.zuoyu;

import org.junit.Before;
import org.junit.Test;
import org.zuoyu.model.PrintVO;
import org.zuoyu.util.ExcelUtil;
import org.zuoyu.util.JsonUtil;

import java.io.InputStream;
import java.util.*;

/**
 * Unit test for simple App.
 */
public class AppTest {
    private final static String PATH = "D:\\GitHub\\esay-excel\\";
    private List<PrintVO> printVOS = new ArrayList<>(16);

    /**
     * Rigorous Test :-)
     */
    @Before
    public void dataLoading() {
        InputStream inputStream = App.class.getClassLoader().getResourceAsStream("data.json");
        List<PrintVO> printVOS = JsonUtil.inputStreamToList(inputStream, PrintVO.class);
        assert printVOS != null;
        this.printVOS.addAll(printVOS);
    }

    @Test
    public void simpleWriteOne() {
        String fileName = PATH + "simpleWriteOne" + System.currentTimeMillis() + ".xlsx";
        ExcelUtil.simpleWrite(fileName, PrintVO.class, printVOS);
    }

    @Test
    public void simpleWriteTwo() {
        String fileName = PATH + "simpleWriteTwo" + System.currentTimeMillis() + ".xlsx";
        ExcelUtil.simpleWrite(fileName, "简单打印2", PrintVO.class, printVOS);
    }

    @Test
    public void simpleWriteThree() {
        String fileName = PATH + "simpleWriteThree" + System.currentTimeMillis() + ".xlsx";
        ExcelUtil.SheetBuilder<PrintVO> sheetBuilder = ExcelUtil.buildSheet("sheet3", PrintVO.class, printVOS);
        ExcelUtil.simpleWrite(fileName, sheetBuilder);
    }

    @Test
    public void simpleWriteFour() {
        String fileName = PATH + "simpleWriteFour" + System.currentTimeMillis() + ".xlsx";
//        只要这些字段
        Set<String> includeColumnFiledNames = new HashSet<>();
        includeColumnFiledNames.add("orderCode");
        includeColumnFiledNames.add("address");
        ExcelUtil.SheetBuilder<PrintVO> sheetBuilder = ExcelUtil.buildSheet("sheet1", PrintVO.class, printVOS, includeColumnFiledNames);
        ExcelUtil.simpleWrite(fileName, sheetBuilder);
    }

    @Test
    public void repeatedWriteOne() {
        String fileName = PATH + "repeatedWriteOne" + System.currentTimeMillis() + ".xlsx";
        ExcelUtil.SheetBuilder<PrintVO> sheetBuilder1 = ExcelUtil.buildSheet( "sheet1", PrintVO.class,printVOS);
        ExcelUtil.SheetBuilder<PrintVO> sheetBuilder2 = ExcelUtil.buildSheet("sheet2",  PrintVO.class,printVOS);
        ExcelUtil.repeatedWrite(fileName, Arrays.asList(sheetBuilder1, sheetBuilder2));
    }

    @Test
    public void repeatedWriteTwo() {
        String fileName = PATH + "repeatedWriteTwo" + System.currentTimeMillis() + ".xlsx";
        ExcelUtil.SheetBuilder<PrintVO> sheetBuilder1 = ExcelUtil.buildSheet( "sheet1", PrintVO.class,printVOS);
        //        只要这些字段
        Set<String> includeColumnFiledNames = new HashSet<>();
        includeColumnFiledNames.add("orderCode");
        includeColumnFiledNames.add("address");
        ExcelUtil.SheetBuilder<PrintVO> sheetBuilder2 = ExcelUtil.buildSheet("sheet2",  PrintVO.class,printVOS, includeColumnFiledNames);
        ExcelUtil.repeatedWrite(fileName, Arrays.asList(sheetBuilder1, sheetBuilder2));
    }
}
