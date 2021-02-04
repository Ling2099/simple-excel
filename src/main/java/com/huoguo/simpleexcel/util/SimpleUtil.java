package com.huoguo.simpleexcel.util;

import com.alibaba.excel.util.StringUtils;

import java.math.BigDecimal;

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
        return !StringUtils.isEmpty(separator) ? str.split(separator) : null;
    }

    /**
     * 类属性值的为空判断
     *
     * @param obj  类属性值
     * @param type 类属性类型
     * @param <T>  表名泛型
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
}
