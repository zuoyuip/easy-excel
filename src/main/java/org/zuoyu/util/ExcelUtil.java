package org.zuoyu.util;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.alibaba.excel.write.style.column.LongestMatchColumnWidthStyleStrategy;
import lombok.AllArgsConstructor;

import java.util.Collection;
import java.util.List;

/**
 * @author : zuoyu
 * @project : esay-excel
 * @description : Excel工具
 * @date : 2019-12-26 10:33
 **/
public class ExcelUtil {


    /**
     * 单Sheet写入
     *
     * @param pathName - 文件位置及文件名
     * @param head     - 头列
     * @param data     - 数据
     * @param <T>      - 类型
     */
    public static <T> void simpleWrite(String pathName, Class<T> head, List<T> data) {
        EasyExcel.write(pathName, head).sheet("sheet1")
                .registerWriteHandler(new LongestMatchColumnWidthStyleStrategy()).doWrite(data);
    }

    /**
     * 单Sheet写入
     *
     * @param pathName  - 文件位置及文件名
     * @param sheetName - sheet名字
     * @param head      - 头列
     * @param data      - 数据
     * @param <T>       - 类型
     */
    public static <T> void simpleWrite(String pathName, String sheetName, Class<T> head, List<T> data) {
        EasyExcel.write(pathName, head).sheet(sheetName)
                .registerWriteHandler(new LongestMatchColumnWidthStyleStrategy()).doWrite(data);
    }

    /**
     * 单Sheet写入
     *
     * @param pathName     - 文件位置及文件名
     * @param sheetBuilder - sheet
     * @param <T>          - 类型
     */
    public static <T> void simpleWrite(String pathName, SheetBuilder<T> sheetBuilder) {
        ExcelWriter excelWriter = EasyExcel.write(pathName).build();
        excelWriter.write(sheetBuilder.data, sheetBuilder.writeSheet);
        excelWriter.finish();
    }

    /**
     * 重复多次写入(多Bean多Sheet)
     *
     * @param pathName      - 文件位置及文件名
     * @param sheetBuilders - sheet
     * @param <T>           - 类型
     */
    public static <T> void repeatedWrite(String pathName, List<SheetBuilder<T>> sheetBuilders) {
        ExcelWriter excelWriter = EasyExcel.write(pathName).build();
        for (int i = 0; i < sheetBuilders.size(); i++) {
            SheetBuilder<T> sheetBuilder = sheetBuilders.get(i);
            sheetBuilder.writeSheet.setSheetNo(i + 1);
            excelWriter.write(sheetBuilder.data, sheetBuilder.writeSheet);
        }
        excelWriter.finish();
    }


    /**
     * 构建Sheet
     *
     * @param sheetName - sheet的名称
     * @param head      - 头列
     * @param data      - 数据
     * @param <T>       - 类型
     * @return - 基于Sheet的构建对象
     */
    public static <T> SheetBuilder<T> buildSheet(String sheetName, Class<T> head, List<T> data) {
        return buildSheet(sheetName, head, data, null);
    }


    /**
     * 构建Sheet
     *
     * @param sheetName               - sheet的名称
     * @param head                    - 头列
     * @param data                    - 数据
     * @param <T>                     - 类型
     * @param includeColumnFiledNames - 选择打印字段（不选择则全部打印）
     * @return - 基于Sheet的构建对象
     */
    public static <T> SheetBuilder<T> buildSheet(String sheetName, Class<T> head, List<T> data, Collection<String> includeColumnFiledNames) {
        WriteSheet writeSheet = EasyExcel.writerSheet(sheetName)
                .includeColumnFiledNames(includeColumnFiledNames)
                .registerWriteHandler(new LongestMatchColumnWidthStyleStrategy())
                .head(head).build();
        return new SheetBuilder<>(data, writeSheet);
    }

    @AllArgsConstructor
    public static class SheetBuilder<T> {
        private List<T> data;
        private WriteSheet writeSheet;
    }
}
