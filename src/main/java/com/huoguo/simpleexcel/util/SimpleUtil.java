package com.huoguo.simpleexcel.util;

import com.alibaba.excel.util.StringUtils;
import com.huoguo.simpleexcel.annotation.ExcelColumns;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.util.*;

/**
 * 简单工具对象
 *
 * @author Lizhenghuang
 */
public final class SimpleUtil {

    /**
     * 字符串过滤,中文符号转英文符号并去空格
     *
     * @param str 输入字符串
     * @return 字符串
     */
    public static String strFilter(String str) {
        return str.replace("！", "!")
                .replace("，", ",")
                .replace("。", "")
                .replace("；", "")
                .replace("：", ":")
                .trim();
    }

    /**
     * 字符串分割
     *
     * @param str       输入字符串
     * @param separator 分隔符
     * @return 字符串数组
     */
    public static String[] strSplit(String str, String separator) {
        return str != null && !"".equals(str) ? str.split(separator) : null;
    }

    /**
     * 转换泛型List集合
     *
     * @param list  数据集合
     * @param clazz 实体类泛型
     * @param <T>   注明泛型
     * @return 实体类集合
     */
    public static <T> List<T> toList(List<Map<Integer, Object>> list, Class<T> clazz, int lineNum) {
        int size = list.size();
        if (size == 0) {
            throw new RuntimeException("The container can not be null");
        }

        List<T> result = new ArrayList<>();

        try {
            for (int i = lineNum; i < size; i++) {
                if (list.get(i).isEmpty()) {
                    throw new RuntimeException("The data cannot be empty");
                }

                T t = clazz.newInstance();
                Field[] fields = t.getClass().getDeclaredFields();

                for (Field field : fields) {
                    field.setAccessible(true);
                    if (field.isAnnotationPresent(ExcelColumns.class)) {
                        ExcelColumns excelColumns = field.getAnnotation(ExcelColumns.class);
                        Object v = list.get(i).get(excelColumns.index());
                        field.set(t, SimpleUtil.convert(v, field.getType()));
                    }
                }
                result.add(t);
            }
        } catch (InstantiationException | IllegalAccessException e) {
            e.printStackTrace();
        }
        return result;
    }

    /**
     * 类属性值的为空判断
     *
     * @param obj  类属性值
     * @param type 类属性类型
     * @param <T>  T
     * @return 泛型对象
     */
    public static <T> T convert(Object obj, Class<T> type) {
        if (obj != null && !StringUtils.isEmpty(obj.toString())) {
            if (type.equals(String.class)) {
                return (T) obj.toString();
            } else if (type.equals(BigDecimal.class)) {
                return (T) new BigDecimal(obj.toString());
            }
        }
        return null;
    }

    /**
     * 获取Excel工作簿
     *
     * @param is 文件输入流
     * @return 工作簿
     */
    public static Workbook getWorkBook(InputStream is) {
        try {
            return WorkbookFactory.create(is);
        } catch (IOException | InvalidFormatException e) {
            e.printStackTrace();
        }
        return null;
    }

    /**
     * 获取工作簿中单元格的值（含合并单元格）
     *
     * @param workbook 当前工作簿
     * @return List
     */
    public static Set<List<String[]>> getWorkbookValue(Workbook workbook) {
        int sheetNum = workbook.getNumberOfSheets();
        Set<List<String[]>> set = new HashSet<>(sheetNum);
        for (int i = 0; i < sheetNum; i++) {
            Sheet sheet = workbook.getSheetAt(i);
            int rows = sheet.getPhysicalNumberOfRows();
            int cells = sheet.getRow(0).getPhysicalNumberOfCells();

            List<String[]> list = new ArrayList<>(rows * cells);
            for (int r = 0; r < rows; r++) {
                Row row = sheet.getRow(r);
                String[] str = new String[cells];

                for (int c = 0; c < cells; c++) {
                    if (isMerged(sheet, r, c)) {
                        str[c] = getMergedValue(sheet, r, c);
                    } else {
                        str[c] = row.getCell(c).toString();
                    }
                }
                list.add(str);
            }
            set.add(list);
        }
        return set;
    }

    /**
     * 判断指定的单元格是否是合并单元格
     *
     * @param sheet  Sheet
     * @param row    行下标
     * @param column 列下标
     * @return true：为合并单元格，反之亦然
     */
    public static boolean isMerged(Sheet sheet, int row, int column) {
        int sheetMergeCount = sheet.getNumMergedRegions();
        for (int i = 0; i < sheetMergeCount; i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            int firstColumn = range.getFirstColumn();
            int lastColumn = range.getLastColumn();
            int firstRow = range.getFirstRow();
            int lastRow = range.getLastRow();
            if (row >= firstRow && row <= lastRow) {
                if (column >= firstColumn && column <= lastColumn) {
                    return true;
                }
            }
        }
        return false;
    }

    /**
     * 获取合并单元格的值
     *
     * @param sheet  Sheet
     * @param row    当前行
     * @param column 当前列
     * @return String
     */
    public static String getMergedValue(Sheet sheet, int row, int column) {
        int sheetMergeCount = sheet.getNumMergedRegions();

        for (int i = 0; i < sheetMergeCount; i++) {
            CellRangeAddress ca = sheet.getMergedRegion(i);
            int firstColumn = ca.getFirstColumn();
            int lastColumn = ca.getLastColumn();
            int firstRow = ca.getFirstRow();
            int lastRow = ca.getLastRow();

            if (row >= firstRow && row <= lastRow) {
                if (column >= firstColumn && column <= lastColumn) {
                    Row fRow = sheet.getRow(firstRow);
                    Cell fCell = fRow.getCell(firstColumn);
                    return getCellValue(fCell);
                }
            }
        }
        return null;
    }

    /**
     * 获取单元格的值
     *
     * @param cell 当前单元格
     * @return String
     */
    public static String getCellValue(Cell cell) {
        return cell == null ? "" : cell.toString();
    }
}
