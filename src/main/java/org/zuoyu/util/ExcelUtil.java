package org.zuoyu.util;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.alibaba.excel.write.style.column.LongestMatchColumnWidthStyleStrategy;
import lombok.AllArgsConstructor;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.net.URLEncoder;
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
        simpleWrite(pathName, "sheet1", head, data);
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
     * Web单Sheet写入
     *
     * @param name      - 文件名
     * @param head      - 头列
     * @param sheetName - 模板名称
     * @param data      - 数据
     * @param <T>       - 类型
     * @param response  - 响应
     */
    public static <T> void simpleWrite(String name, Class<T> head, String sheetName, List<T> data, HttpServletResponse response) throws IOException {

        try {
            response.setContentType("application/vnd.ms-excel");
            response.setCharacterEncoding("utf-8");
            String fileName = URLEncoder.encode(name, "UTF-8");
            response.setHeader("Content-disposition", "attachment;filename=" + fileName + ".xlsx");
            EasyExcel.write(response.getOutputStream(), head).autoCloseStream(Boolean.FALSE).sheet(sheetName).registerWriteHandler(new LongestMatchColumnWidthStyleStrategy()).doWrite(data);
        } catch (Exception e) {
            response.reset();
            response.setContentType("application/json");
            response.setCharacterEncoding("UTF-8");
            response.sendError(HttpServletResponse.SC_INTERNAL_SERVER_ERROR, "下载文件失败");
        }
    }

    /**
     * Web单Sheet写入
     *
     * @param name     - 文件名
     * @param head     - 头列
     * @param data     - 数据
     * @param <T>      - 类型
     * @param response - 响应
     */
    public static <T> void simpleWrite(String name, Class<T> head, List<T> data, HttpServletResponse response) throws IOException {
        simpleWrite(name, head, "sheet1", data, response);
    }

    /**
     * Web单Sheet写入
     *
     * @param name     - 文件名
     * @param sheetBuilder - sheet
     * @param response - 响应
     * @param <T>          - 类型
     */
    public static <T> void simpleWrite(String name, SheetBuilder<T> sheetBuilder, HttpServletResponse response) throws IOException {
        try {
            response.setContentType("application/vnd.ms-excel");
            response.setCharacterEncoding("utf-8");
            String fileName = URLEncoder.encode(name, "UTF-8");
            response.setHeader("Content-disposition", "attachment;filename=" + fileName + ".xlsx");
            ExcelWriter excelWriter = EasyExcel.write(response.getOutputStream()).autoCloseStream(Boolean.FALSE).build();
            excelWriter.write(sheetBuilder.data, sheetBuilder.writeSheet);
            excelWriter.finish();
        } catch (Exception e) {
            response.reset();
            response.setContentType("application/json");
            response.setCharacterEncoding("UTF-8");
            response.sendError(HttpServletResponse.SC_INTERNAL_SERVER_ERROR, "下载文件失败");
        }
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
        repeatedWriter(excelWriter, sheetBuilders);
    }

    /**
     * Web重复多次写入(多Bean多Sheet)
     *
     * @param name          - 文件名
     * @param sheetBuilders - sheet
     * @param response      - 响应
     * @param <T>           - 类型
     */
    public static <T> void repeatedWrite(String name, List<SheetBuilder<T>> sheetBuilders, HttpServletResponse response) throws IOException {
        try {
            response.setContentType("application/vnd.ms-excel");
            response.setCharacterEncoding("utf-8");
            String fileName = URLEncoder.encode(name, "UTF-8");
            response.setHeader("Content-disposition", "attachment;filename=" + fileName + ".xlsx");
            ExcelWriter excelWriter = EasyExcel.write(response.getOutputStream()).autoCloseStream(Boolean.FALSE).build();
            repeatedWriter(excelWriter, sheetBuilders);
        } catch (Exception e) {
            response.reset();
            response.setContentType("application/json");
            response.setCharacterEncoding("UTF-8");
            response.sendError(HttpServletResponse.SC_INTERNAL_SERVER_ERROR, "下载文件失败");
        }
    }

    private static <T> void repeatedWriter(ExcelWriter excelWriter, List<SheetBuilder<T>> sheetBuilders) {
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
        return buildSheet(sheetName, head, data, includeColumnFiledNames, null);
    }

    /**
     * 构建Sheet
     *
     * @param sheetName               - sheet的名称
     * @param head                    - 头列
     * @param data                    - 数据
     * @param <T>                     - 类型
     * @param includeColumnFiledNames - 选择打印字段（不选择则全部打印）
     * @param excludeColumnFiledNames - 选择忽略打印的字段（不选择则全部打印）
     * @return - 基于Sheet的构建对象
     */
    public static <T> SheetBuilder<T> buildSheet(String sheetName, Class<T> head, List<T> data, Collection<String> includeColumnFiledNames, Collection<String> excludeColumnFiledNames) {
        WriteSheet writeSheet = EasyExcel.writerSheet(sheetName)
                .includeColumnFiledNames(includeColumnFiledNames)
                .excludeColumnFiledNames(excludeColumnFiledNames)
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
