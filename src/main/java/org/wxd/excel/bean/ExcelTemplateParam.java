package org.wxd.excel.bean;

import com.google.common.collect.Maps;
import org.wxd.excel.annotation.ExcelParam;
import org.wxd.excel.bean.handler.ExcelTemplateRedRepositoryHandler;
import org.wxd.excel.exception.ExcelException;
import org.wxd.excel.utils.Assert;

import java.lang.reflect.Field;
import java.util.Arrays;
import java.util.List;
import java.util.Map;

/**
 * @Description:
 * @Author : wangxd
 * @Date: 2016-3-3
 * @Version 1.0
 */
public class ExcelTemplateParam {
    private String sheetTitle;
    private Map<String,Object> params = Maps.newConcurrentMap();

    public static Builder newBuilder(){
        return new Builder();
    }
    ExcelTemplateParam(Builder builder){
        this.sheetTitle = builder.sheetTitle;
        this.params = builder.params;
    }

    public static class Builder implements ExcelTemplateRedRepositoryHandler<Builder> {
        String sheetTitle;
        Map<String,Object> params = Maps.newConcurrentMap();


        @Override
        public Builder readExcelRepository(ExcelRepository repository) {
            Assert.notNull(repository, "repository can not be null.");
            try {
                if(repository.excelSheet() == null) return this;
                this.sheetTitle = repository.excelSheet().name();
                /*读取本身属性*/
                this.reflectFieldToMap(repository.fields(),repository.target());
                /*父类属性*/
                this.reflectFieldToMap(Arrays.asList(repository.targetClass().getSuperclass().getDeclaredFields()),repository.target());
                return this;
            } catch (Exception e) {
                throw new ExcelException(e.getMessage(), e);
            }
        }
        /**
         * 反射读取属性
         * @param fields
         * @param target
         */
        private void reflectFieldToMap(List<Field> fields,Object target){
            try{
                for (Field field : fields) {
                    field.setAccessible(true);
                    if (!field.isAnnotationPresent(ExcelParam.class)) continue;
                    ExcelParam excelParam = field.getAnnotation(ExcelParam.class);
                    /*读取参数注解*/
                    this.params.put(excelParam.name(),field.get(target) == null ? "" : field.get(target));
                }
            }catch (Exception e){
                throw new ExcelException(e.getMessage(),e);
            }
        }
        public ExcelTemplateParam build(){
            return new ExcelTemplateParam(this);
        }

        public Builder sheetTitle(String sheetTitle) {
            this.sheetTitle = sheetTitle;
            return this;
        }

        public Builder params(Map<String, Object> params) {
            this.params = params;
            return this;
        }
    }

    public String sheetTitle() {
        return sheetTitle;
    }

    public Map<String, Object> params() {
        return params;
    }
}
