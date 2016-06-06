package org.wxd.excel.bean;

import com.google.common.collect.Lists;

import java.util.List;

/**
 * @Description: excel总体内容
 * @Author : wangxd
 * @Date: 2016-3-3
 * @Version 1.0
 */
public class ExcelContent{

    private List<ExcelTemplate> templates;/*存放类上的ExcelCell注解*/
    private List<ExcelTemplateParam> params;/*存放类上的ExcelParam注解*/
    private List<ExcelTemplateFormula> formulas;/*存放类上的ExcelFormula注解*/


    ExcelContent(Builder builder) {
        this.templates = builder.templates;
        this.params = builder.params;
        this.formulas = builder.formulas;
    }

    public static Builder newBuilder(){
        return new Builder();
    }

    public static class Builder{
        List<ExcelTemplate> templates = Lists.newArrayList();/*存放类上的ExcelCell注解*/
        List<ExcelTemplateParam> params = Lists.newArrayList();/*存放类上的ExcelParam注解*/
        List<ExcelTemplateFormula> formulas = Lists.newArrayList();/*存放类上的ExcelFormula注解*/

        public ExcelContent build(){
            return new ExcelContent(this);
        }

        public Builder addTemplate(ExcelTemplate template) {
            this.templates.add(template);
            return this;
        }

        public Builder addParam(ExcelTemplateParam param) {
            this.params.add(param);
            return this;
        }

        public Builder addFormula(ExcelTemplateFormula formula) {
            this.formulas.add(formula);
            return this;
        }
    }

    public List<ExcelTemplateFormula> formulas() {
        return formulas;
    }

    public List<ExcelTemplate> templates() {
        return templates;
    }

    public List<ExcelTemplateParam> params() {
        return params;
    }

    @Override
    public String toString() {
        return "ExcelContent{" +
                "templates=" + templates +
                ", params=" + params +
                ", formulas=" + formulas +
                '}';
    }
}
