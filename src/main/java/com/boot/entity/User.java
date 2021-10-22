package com.boot.entity;

import com.boot.annonation.ExcelFiled;
import com.boot.enums.ExcelFiledType;
import lombok.*;

import java.time.LocalDateTime;
import java.time.LocalTime;

/**
 * @Desc 用户 - 实体
 * @Author lixu
 * @Date 2021/10/14 11:22
 */
@Setter
@Getter
@AllArgsConstructor
@NoArgsConstructor
@Builder
public class User {

    /**
     * PK
     */
    @ExcelFiled(title = "PK", columnIndex = 0, rowIndex = 1, sort = 0, type = ExcelFiledType.INT)
    private Long id;

    /**
     * 名称
     */
    @ExcelFiled(title = "名称", columnIndex = 1, rowIndex = 1, sort = 1)
    private String userName;

    /**
     * 性别
     */
    @ExcelFiled(title = "性别", columnIndex = 2, rowIndex = 1, sort = 2, type = ExcelFiledType.INT)
    private Integer sex;

    /**
     * 出生日期
     */
    @ExcelFiled(title = "出生日期", columnIndex = 3, rowIndex = 1, sort = 3)
    private String birthday;

    /**
     * 用来测试是否受到非注解字段影响
     */
    private String testField;
}
