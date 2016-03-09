package org.wxd.excel.utils;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.wxd.excel.annotation.Excel;
import org.wxd.excel.bean.ExcelContent;
import org.wxd.excel.bean.ExcelRepository;
import org.wxd.excel.bean.ExcelTemplate;
import org.wxd.excel.bean.ExcelTemplateParam;
import org.wxd.excel.exception.ExcelException;
import org.wxd.excel.handler.EntityHandlerExecutor;
import org.wxd.excel.handler.ExcelHandlerExecutor;

import java.io.File;
import java.lang.reflect.Field;
import java.util.List;

/**
 * @Description:
 * @Author : wangxd
 * @Date: 2016-3-2
 * @Version 1.0
 */
public class ExportFatory {

    /**
     * 根据模板创建Workbook
     * @param excelFile
     * @return
     */
    public static Workbook buildWorkbookOfTemplate(File excelFile) {
        try {
            return WorkbookFactory.create(excelFile);
        } catch (Exception e) {
            throw new ExcelException(e.getMessage(), e);
        }
    }
    /**
     *创建excel处理注册器
     * @return
     */
    public static ExcelHandlerExecutor buildExecutor(){
        return ExcelHandlerExecutor.instance();
    }

    /**
     * 创建excel转实体注册器
     * @return
     */
    public static EntityHandlerExecutor buildEntityExecutor(){return EntityHandlerExecutor.instance();}

    /**
     * TODO 待解决字节码解析问题
     * 将实体信息存于ExcelContent中
     * @param src
     * @return
     */
    public static ExcelContent buildExcelContent(Object src){
        Assert.notNull(src, "src can not be null.");
        ExcelContent content = ExcelContent.newBuilder().build();
        try{
            Field[] fields = src.getClass().getDeclaredFields();
            for (Field field : fields) {
                if(!field.isAnnotationPresent(Excel.class)) continue;
                field.setAccessible(true);
                Object targetFieldValue = field.get(src);
                if(targetFieldValue == null) continue;
                Field[] propFields = targetFieldValue.getClass().getDeclaredFields();
                for (Field propField : propFields) {
                    propField.setAccessible(true);
                    if(!(propField.get(targetFieldValue) instanceof List)) continue;
                    List list = (List) propField.get(targetFieldValue);
                    for (Object targetObj : list) {
                        ExcelRepository excelRepository = ExcelRepository.newBuilder().read(targetObj).build();
                        ExcelTemplate excelTemplate = ExcelTemplate.newBuilder().readExcelRepository(excelRepository).build();
                        ExcelTemplateParam excelTemplateParam = ExcelTemplateParam.newBuilder().readExcelRepository(excelRepository).build();
                        content.templates().add(excelTemplate);
                        content.params().add(excelTemplateParam);
                    }
                }
            }
            return content;
        }catch (Exception e){
            throw new ExcelException("src transfore input ExcelContent error.",e);
        }
    }

}
