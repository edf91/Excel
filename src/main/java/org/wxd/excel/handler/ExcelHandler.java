package org.wxd.excel.handler;

import org.apache.poi.ss.usermodel.Workbook;
import org.wxd.excel.bean.ExcelContent;

/**
 * @Description: excel处理handler，用户可以实现该接口，并注入实现自定义excel处理
 * @Author : wangxd
 * @Date: 2016-3-3
 * @Version 1.0
 */
public interface ExcelHandler {
    /**
     * 处理excel接口
     * @param workbook 目标excel
     * @param content 待填充内容
     * @param custom 用户自定义对象
     * @return
     */
    Workbook handlerWorkbook(Workbook workbook, ExcelContent content, Object custom);
}
