package com.boot.enums;

import lombok.AllArgsConstructor;
import lombok.Getter;

/**
 * @Desc Excel 导出字段类型
 * @Author lixu
 * @Date 2021/10/22 13:55
 */
@Getter
@AllArgsConstructor
public enum ExcelFiledType {

    /**
     * 类型
     */
    STRING,
    DOUBLE,
    INT;
}
