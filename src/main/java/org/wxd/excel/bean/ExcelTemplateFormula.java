package org.wxd.excel.bean;

import com.google.common.collect.Lists;
import org.wxd.excel.annotation.ExcelFormula;
import org.wxd.excel.bean.handler.ExcelTemplateRedRepositoryHandler;
import org.wxd.excel.exception.ExcelException;
import org.wxd.excel.utils.Assert;

import java.lang.reflect.Field;
import java.util.Arrays;
import java.util.List;

/**
 * @Description: 单元格计算公式
 * @Copyright: Copyright 2012 ShenZhen DSE Corporation
 * @Company: 深圳市东深电子股份有限公司
 * @Author : wangxd
 * @Date: 2016-3-11
 * @Version 1.0
 */
public class ExcelTemplateFormula {
    private String sheetTitle;
    private List<ExcelFormula> formulas;

    ExcelTemplateFormula(Builder builder) {
        this.sheetTitle = builder.sheetTitle;
        this.formulas = builder.formulas;
    }


    public static Builder newBuilder(){
        return new Builder();
    }
    public static  class Builder implements ExcelTemplateRedRepositoryHandler<Builder> {
        String sheetTitle;
        List<ExcelFormula> formulas = Lists.newArrayList();

        public ExcelTemplateFormula build(){
            return new ExcelTemplateFormula(this);
        }
        /**
         * 读取ExcelRepository对象，拿取里面的数据
         * @param repository
         * @return
         */
        @Override
        public Builder readExcelRepository(ExcelRepository repository) {
            Assert.notNull(repository, "repository can not be null.");
            try {
                if (repository.excelSheet() == null) return this;
                this.sheetTitle = repository.excelSheet().name();
                /*读取本身属性*/
                this.reflectFieldToMap(repository.fields());
                /*父类属性*/
                this.reflectFieldToMap(Arrays.asList(repository.targetClass().getSuperclass().getDeclaredFields()));
                return this;
            } catch (Exception e) {
                throw new ExcelException(e.getMessage(), e);
            }
        }

        /**
         * 反射读取属性
         * @param fields
         */
        private void reflectFieldToMap(List<Field> fields){
            try{
                for (Field field : fields) {
                    field.setAccessible(true);
                    if (!field.isAnnotationPresent(ExcelFormula.class)) continue;
                    this.formulas.add(field.getAnnotation(ExcelFormula.class));
                }
            }catch (Exception e){
                throw new ExcelException(e.getMessage(),e);
            }
        }

        public Builder sheetTitle(String sheetTitle) {
            this.sheetTitle = sheetTitle;
            return this;
        }

        public Builder formulas(List<ExcelFormula> formulas) {
            this.formulas = formulas;
            return this;
        }
    }

    public String sheetTitle() {
        return sheetTitle;
    }

    public List<ExcelFormula> formulas() {
        return formulas;
    }

    @Override
    public String toString() {
        return "ExcelTemplateFormula{" +
                "sheetTitle='" + sheetTitle + '\'' +
                ", formulas=" + formulas +
                '}';
    }
}
