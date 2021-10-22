package com.boot.annonation;

import com.boot.enums.ExcelFiledType;

import java.lang.annotation.*;

/**
 * @Desc Excel导入导出字段注解
 * @Author lixu
 * @Date 2021/10/14 11:18
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
@Documented
public @interface ExcelFiled {

    /**
     * 列标题
     */
    String title() default "";

    /**
     * 列开始下标
     */
    int columnIndex() default 0;

    /**
     * 行开始下标
     */
    int rowIndex() default 0;

    /**
     * 导出 Excel 字段排序
     */
    int sort() default 0;

    /**
     * 导出 Excel 类型
     */
    ExcelFiledType type() default ExcelFiledType.STRING;
}
