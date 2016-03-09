package org.wxd.excel.bean;

import com.google.common.collect.Lists;
import org.wxd.excel.annotation.ExcelSheet;
import org.wxd.excel.exception.ExcelException;
import org.wxd.excel.utils.Assert;

import java.lang.reflect.Field;
import java.util.Arrays;
import java.util.List;

/**
 * @Description: 存放对象有关于Excel的所有信息
 * @Author : wangxd
 * @Date: 2016-3-2
 * @Version 1.0
 */
public final class ExcelRepository {

    private Object target;/*目标对象*/
    private ExcelSheet excelSheet;/*sheet注解*/
    private Class targetObjClass;/*目标类字节码*/
    private Class targetClass;/*目标对象字节码*/
    private List<Field> fields = Lists.newArrayList();/*属性集合*/


    ExcelRepository(Builder builder) {
        this.target = builder.target;
        this.excelSheet = builder.excelSheet;
        this.targetObjClass = builder.targetObjClass;
        this.targetClass = builder.targetClass;
        this.fields = builder.fields;
    }

    public static ExcelRepository.Builder newBuilder() {
        return new Builder();
    }

    public  static class Builder {
        Object target;/*目标对象*/
        ExcelSheet excelSheet;/*sheet注解*/
        Class targetObjClass;/*目标类字节码*/
        Class targetClass;/*目标对象字节码*/
        List<Field> fields = Lists.newArrayList();/*属性集合*/

        Builder() {
        }

        public Builder read(Object target) {
            try {
                Assert.notNull(target, "target can not be null.");
                this.targetObjClass = Class.forName(target.getClass().getName());
                if (!this.targetObjClass.isAnnotationPresent(ExcelSheet.class)) return this;
                this.target = target;
                this.excelSheet = (ExcelSheet) this.targetObjClass.getAnnotation(ExcelSheet.class);
                this.targetClass = target.getClass();
                this.fields = Arrays.asList(this.targetObjClass.getDeclaredFields());
                return this;
            } catch (Exception e) {
                throw new ExcelException(e.getMessage(), e);
            }
        }

        public ExcelRepository build() {
            return new ExcelRepository(this);
        }

        public Builder target(Object target) {
            this.target = target;
            return this;
        }

        public Builder excelSheet(ExcelSheet excelSheet) {
            this.excelSheet = excelSheet;
            return this;
        }


        public Builder targetObjClass(Class targetObjClass) {
            this.targetObjClass = targetObjClass;
            return this;
        }

        public Builder targetClass(Class targetClass) {
            this.targetClass = targetClass;
            return this;
        }

        public Builder fields(List<Field> fields) {
            this.fields = fields;
            return this;
        }
    }

    public Object target() {
        return target;
    }

    public ExcelSheet excelSheet() {
        return excelSheet;
    }


    public Class targetObjClass() {
        return targetObjClass;
    }

    public Class targetClass() {
        return targetClass;
    }

    public List<Field> fields() {
        return fields;
    }

}
