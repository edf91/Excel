package org.wxd.excel.handler.inport;

import org.apache.poi.ss.usermodel.Workbook;

import java.util.List;
import java.util.Map;

/**
 * @Description:
 * @Author : wangxd
 * @Date: 2016-3-7
 * @Version 1.0
 */
public interface EntityHandler {

    /**
     * 处理excel转实体接口，返回一个map，sheetTitle:对象集合
     * @param workbook 目标Workbook
     * @param result 处理结果
     * @param custom 自定义对象，满足其他需求
     * @return
     */
    public Map<String,List<Object>> handlerExcelToEntity(Workbook workbook, Map<String, List<Object>> result, Object custom);

}
