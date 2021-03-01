package com.huoguo.simpleexcel.service;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.EasyExcelFactory;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.alibaba.excel.write.metadata.fill.FillConfig;
import com.huoguo.simpleexcel.util.SimpleUtil;
import org.apache.poi.ss.usermodel.Workbook;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.InputStream;
import java.net.URLEncoder;
import java.util.List;
import java.util.Map;


/**
 * Excel工具接口实现类
 *
 * @author Lizhenghuang
 */
public class ExcelService {

    /**
     * 解析Excel
     *
     * @param is  文件输入流
     * @param <T> 注明泛型
     * @return List
     */
    public static <T> List<T> importData(InputStream is) {
        return EasyExcelFactory.read(is).doReadAllSync();
    }

    /**
     * 解析Excel
     *
     * @param is    文件输入流
     * @param clazz 实体类
     * @param <T>   注明泛型
     * @return 实体类集合
     */
    public static <T> List<T> importData(InputStream is, Class<T> clazz) {
        List<Map<Integer, Object>> list = EasyExcelFactory.read(is).doReadAllSync();
        return SimpleUtil.toList(list, clazz, 0);
    }

    /**
     * 解析Excel
     *
     * @param is      文件输入流
     * @param clazz   实体类
     * @param lineNum 固定行开始解析
     * @param <T>     注明泛型
     * @return 实体类集合
     */
    public static <T> List<T> importData(InputStream is, Class<T> clazz, int lineNum) {
        List<Map<Integer, Object>> list = EasyExcelFactory.read(is).doReadAllSync();
        return SimpleUtil.toList(list, clazz, lineNum);
    }

    /**
     * 解析Excel（含合并单元格）
     *
     * @param is 文件输入流
     * @return List
     */
    public static List<String[]> analyseData(InputStream is) {
        Workbook workbook = SimpleUtil.getWorkBook(is);
        if (workbook == null) {
            throw new RuntimeException("This Workbook is Null");
        }
        return SimpleUtil.getWorkbookValue(workbook);
    }

    /**
     * 填充类型的文件下载，需预先定义好模板
     *
     * @param response Http响应
     * @param request  Http请求
     * @param is       文件输入流
     * @param list     数据集合
     * @param fileName 下载后文件的名称
     * @param map      附加选项（填充）
     */
    public static void exportData(HttpServletResponse response, HttpServletRequest request, InputStream is, List list, String fileName, Map<String, Object> map) {
        try {
            response.setCharacterEncoding("utf-8");
            response.setHeader("Content-Disposition", "attachment; filename=" + URLEncoder.encode(fileName, "UTF-8"));
            response.setHeader("fileName", fileName);
            response.setHeader("Access-control-Allow-Origin", request.getHeader("Origin"));
            response.setHeader("Pragma", "public");
            response.setHeader("Cache-Control", "no-store");
            response.addHeader("Cache-Control", "max-age=0");

            if (map != null) {
                ExcelWriter excelWriter = EasyExcel.write(response.getOutputStream()).withTemplate(is).build();
                WriteSheet writeSheet = EasyExcel.writerSheet().build();
                // 设置List在向下填充时，自动添加一行空白
                FillConfig fillConfig = FillConfig.builder().forceNewRow(Boolean.TRUE).build();
                excelWriter.fill(list, fillConfig, writeSheet);
                excelWriter.fill(map, writeSheet);
                excelWriter.finish();
            } else {
                EasyExcel.write(response.getOutputStream()).withTemplate(is).sheet().doFill(list);
            }

            response.flushBuffer();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * 简单的文件流下载，配合easyExcel实体类的注解一起使用
     * 例如：@ExcelProperty("字符串标题") 或 @ExcelIgnore
     *
     * @param response Http响应
     * @param list     数据集合
     * @param fileName 文件名称
     */
    public static void exportData(HttpServletResponse response, List list, String fileName) {
        try {
            response.setCharacterEncoding("utf-8");
            response.setHeader("Content-Disposition", "attachment; filename=" + URLEncoder.encode(fileName, "UTF-8"));
            response.setHeader("Pragma", "public");
            response.setHeader("Cache-Control", "no-store");
            response.addHeader("Cache-Control", "max-age=0");

            EasyExcel.write(response.getOutputStream()).sheet().doWrite(list);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * 简单的文件流下载，配合easyExcel实体类的注解一起使用
     * 例如：@ExcelProperty("字符串标题") 或 @ExcelIgnore
     *
     * @param response Http响应
     * @param list     数据集合
     * @param fileName 文件名称
     * @param clazz    文件头的实体类
     */
    public static void exportData(HttpServletResponse response, List list, String fileName, Class clazz) {
        try {
            response.setCharacterEncoding("utf-8");
            response.setHeader("Content-Disposition", "attachment; filename=" + URLEncoder.encode(fileName, "UTF-8"));
            response.setHeader("Pragma", "public");
            response.setHeader("Cache-Control", "no-store");
            response.addHeader("Cache-Control", "max-age=0");

            EasyExcel.write(response.getOutputStream(), clazz).sheet().doWrite(list);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
