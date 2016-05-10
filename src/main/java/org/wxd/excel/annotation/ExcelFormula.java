package org.wxd.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * @Description: Excel公式注解
 * @Copyright: Copyright 2012 ShenZhen DSE Corporation
 * @Company: 深圳市东深电子股份有限公司
 * @Author : wangxd
 * @Date: 2016-3-11
 * @Version 1.0
 */
@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.FIELD})
public @interface ExcelFormula {
    /**
     * 单元格位置，或者起始单元格位置：
     * 列：行
     * @return
     */
    String index();

    /**
     * 是否其后的该行所有单元格都使用该公式；即编辑Excel时横向拉应用的效果
     * @return
     */
    boolean isApplyColum() default false;

    /**
     * 是否行列相应自增
     * @return
     */
    boolean isIncrement() default false;

    /**
     * 是否其后的该列所有单元格都是用该公式；即编辑Excel时下拉应用的效果
     * @return
     */
    boolean isApplyRow() default false;

    /**
     * Excel计算公式
     * *代表之后的所有都加入公式，如SUM(B6:B*),代表B6到B整列
     * @return
     */
    String formula();


}
