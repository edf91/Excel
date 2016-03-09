package org.wxd.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * @Description: 类上加上该注解，表示该类对应Excel中的一个Sheet
 * @Author : wangxd
 * @Date: 2016-3-2
 * @Version 1.0
 */
@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.TYPE})
public @interface ExcelSheet {
    /**
     * Sheet名称
     * @return
     */
    String name() default "";

    /**
     * 开始输出的行
     * @return
     */
    int beginWriteRowIndex() default 0;

    int beginReadRowIndex() default 0;
}
