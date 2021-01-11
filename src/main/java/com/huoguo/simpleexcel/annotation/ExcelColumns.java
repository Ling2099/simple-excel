package com.huoguo.simpleexcel.annotation;

import java.lang.annotation.*;

/**
 * Excel解析实体类类属性注解
 * @author Lizhenghuang
 */
@Documented
@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.FIELD, ElementType.ANNOTATION_TYPE})
public @interface ExcelColumns {
    /**
     * 序号
     * @return int
     */
    int index();
}
