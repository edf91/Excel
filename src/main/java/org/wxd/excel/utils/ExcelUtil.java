package org.wxd.excel.utils;

import org.apache.poi.ss.usermodel.Workbook;
import org.wxd.excel.handler.inport.ExcelHandler;

import java.io.File;
import java.util.Arrays;
import java.util.List;

/**
 * @Description:
 * @Copyright: Copyright 2012 ShenZhen DSE Corporation
 * @Company: 深圳市东深电子股份有限公司
 * @Author : wangxd
 * @Date: 2016-5-10
 * @Version 1.0
 */
public class ExcelUtil {


    /**
     * 将实体根据模板读取到workbook
     * @param src 目标对象
     * @param sheetTiltles 保留的sheet
     * @param excelTemplateFile 模板文件
     * @param handlers 定义的处理器
     * @return
     */
    public static Workbook buildWorkbookFromEntityOfFile(Object src,List<String> sheetTiltles, File excelTemplateFile,ExcelHandler... handlers) {
        try {
            Workbook workbook = ExportFatory.buildWorkbookOfTemplate(excelTemplateFile);
            ExportFatory.buildExecutor().registerAll(Arrays.asList(handlers)).handler(
                    workbook,
                    ExportFatory.buildExcelContent(src),
                    sheetTiltles
            );
            return workbook;
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }
}
