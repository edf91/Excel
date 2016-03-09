package org.wxd.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * @Description: ExcelCell单元格注解
 * @Author : wangxd
 * @Date: 2016-3-2
 * @Version 1.0
 */
@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.FIELD})
public @interface ExcelCell {
    /**
     * 单元格在一行里的序号,越小越前
     * @return
     */
    int order() default -1;

    /**
     * 单元格样式
     * @return
     */
    ExcelCellStyle[] style() default ExcelCellStyle.NULL;

}
