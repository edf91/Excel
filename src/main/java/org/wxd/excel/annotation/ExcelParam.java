package org.wxd.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * @Description:  占位符${属性名称}，查找到改字符并替换值
 * @Author : wangxd
 * @Date: 2016-3-3
 * @Version 1.0
 */
@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.FIELD})
public @interface ExcelParam {
    String name() default "";

    /**
     * 读取的位置:（行数:列数）
     * @return
     */
    String readIndex() default "";
}
