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

    private List<ExcelTemplate> templates = Lists.newArrayList();/*存放类上的ExcelCell注解*/
    private List<ExcelTemplateParam> params = Lists.newArrayList();/*存放类上的ExcelParam注解*/

    ExcelContent(Builder builder) {
        this.templates = builder.templates;
        this.templates = builder.templates;
    }

    public static Builder newBuilder(){
        return new Builder();
    }

    public static class Builder{
        List<ExcelTemplate> templates = Lists.newArrayList();/*存放类上的ExcelCell注解*/
        List<ExcelTemplateParam> params = Lists.newArrayList();/*存放类上的ExcelParam注解*/

        public Builder templates(List<ExcelTemplate> templates) {
            this.templates = templates;
            return this;
        }

        public Builder params(List<ExcelTemplateParam> params) {
            this.params = params;
            return this;
        }

        public ExcelContent build(){
            return new ExcelContent(this);
        }

    }

    public List<ExcelTemplate> templates() {
        return templates;
    }

    public List<ExcelTemplateParam> params() {
        return params;
    }
}
