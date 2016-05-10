package org.wxd.excel.handler.export;

import com.google.common.collect.Maps;
import org.apache.poi.ss.usermodel.*;
import org.wxd.excel.annotation.ExcelCellStyle;
import org.wxd.excel.bean.CellInfo;
import org.wxd.excel.bean.ExcelContent;
import org.wxd.excel.bean.ExcelTemplate;
import org.wxd.excel.bean.ExcelTemplateParam;
import org.wxd.excel.handler.inport.ExcelHandler;

import java.math.BigDecimal;
import java.util.List;
import java.util.Map;

/**
 * @Description: 默认导出处理器
 * @Copyright: Copyright 2012 ShenZhen DSE Corporation
 * @Company: 深圳市东深电子股份有限公司
 * @Author : wangxd
 * @Date: 2016-5-10
 * @Version 1.0
 */
public class DefaultExportHandler implements ExcelHandler {

    @Override
    public Workbook handlerWorkbook(Workbook workbook, ExcelContent content,Object custom) {

        List<String> sheetTitles = (List<String>) custom;
        List<ExcelTemplate> excelTemplates = content.templates();
        List<ExcelTemplateParam> excelTemplateParams = content.params();

        Sheet sheet = null;
        Row row = null;
        Cell cell = null;
        CellStyle style = workbook.createCellStyle();

        Map<String,Boolean> isNeedToRemoveSheet = Maps.newHashMap();
        /*移除不必要的sheet*/
        for(int i = 0,len = workbook.getNumberOfSheets(); i < len; i++){
            String sheetName = workbook.getSheetAt(i).getSheetName();
            if (!sheetTitles.contains(sheetName)) {
                workbook.removeSheetAt(i);
                len = workbook.getNumberOfSheets();
                i --;
                continue;
            }
            isNeedToRemoveSheet.put(sheetName,Boolean.TRUE);
        }




            /*处理参数*/
        for (ExcelTemplateParam excelTemplateParam : excelTemplateParams) {
            if(excelTemplateParam.sheetTitle() == null || "".equals(excelTemplateParam.sheetTitle())) continue;
            sheet = workbook.getSheet(excelTemplateParam.sheetTitle());
            isNeedToRemoveSheet.remove(excelTemplateParam.sheetTitle());
                /*处理参数*/
            for (int index = sheet.getFirstRowNum(); index <= sheet.getLastRowNum(); index++) {
                row = sheet.getRow(index);
                for (int cellIndex = row.getFirstCellNum(); cellIndex <= row.getLastCellNum(); cellIndex++) {
                    if (cellIndex < 0) break;
                    cell = row.getCell(cellIndex);
                    if (cell == null) continue;
                    String value = cell.getStringCellValue();
                    if (!value.contains("${")) continue;
                    int valueLength = value.length();
                    int subLength = 0;
                    for (int i = value.indexOf("$"); i < valueLength; i++) {
                        if (value.charAt(i) != '$') continue;
                        int begIndex = i;
                        for (int j = i; j < valueLength; j++) {
                            i++;
                            if (value.charAt(j) != '}') continue;
                            String paramName = value.substring(begIndex + 2, j);
                            String paramValue = excelTemplateParam.params().get(paramName) == null ? "" : excelTemplateParam.params().get(paramName).toString(); //excelTemplate.params().get(paramName) == null ? "" : excelTemplate.params().get(paramName).toString();

                            value = value.replace("${" + paramName + "}", paramValue);
                            subLength = Math.abs(value.length() - valueLength);
                            valueLength = value.length();
                            i = Math.abs(i - subLength);
                            break;
                        }
                        if (i == valueLength || i > valueLength) break;
                    }
                    cell.setCellValue(value);
                }
            }
        }

        Map<String, Integer> hasDealIndexMap = Maps.newHashMap();
        for (ExcelTemplate excelTemplate : excelTemplates) {
            if (excelTemplate.sheetTitle() == null) continue;
            Integer currentIndex = hasDealIndexMap.get(excelTemplate.sheetTitle()) == null ? excelTemplate.beginWriteRowIndex() : hasDealIndexMap.get(excelTemplate.sheetTitle());
            sheet = workbook.getSheet(excelTemplate.sheetTitle());

            sheet.shiftRows(currentIndex, sheet.getLastRowNum(), 1, true, false);
            row = sheet.createRow(currentIndex);
            row.setHeight((short) (20 * 18));
            for (Map.Entry<Integer, CellInfo> entry : excelTemplate.orderCellMap().entrySet()) {
                CellInfo cellInfo = entry.getValue();
                String cellValue = cellInfo.value() == null ? "" : cellInfo.value().toString();
                if (cellInfo.order() == -1) continue;
                cell = row.createCell(cellInfo.order());
                if(cellInfo.fieldType().toString().equals("class java.math.BigDecimal")){
                    cell.setCellValue(new BigDecimal(cellValue.equals("") ?  "0" : cellValue).doubleValue());
                }else{
                    cell.setCellValue(cellValue);
                }
                if (cellInfo.styles() == null) continue;
                for (ExcelCellStyle excelCellStyle : cellInfo.styles()) {
                    if (excelCellStyle.equals(ExcelCellStyle.BORDER_ALL)) {
                        style.setBorderBottom(CellStyle.BORDER_THIN); //下边框
                        style.setBorderLeft(CellStyle.BORDER_THIN);//左边框
                        style.setBorderTop(CellStyle.BORDER_THIN);//上边框
                        style.setBorderRight(CellStyle.BORDER_THIN);//右边框
                    }
                    if (excelCellStyle.equals(ExcelCellStyle.ALIGN_CENTER)) {
                        style.setAlignment(CellStyle.ALIGN_CENTER);
                    }
                }
                cell.setCellStyle(style);
            }
            hasDealIndexMap.put(excelTemplate.sheetTitle(), ++currentIndex);
        }

        /*删除没有数据的sheet*/
       /* if(!isNeedToRemoveSheet.isEmpty()){
            for(int i = 0,len = workbook.getNumberOfSheets(); i < len; i++){
                String sheetName = workbook.getSheetAt(i).getSheetName();
                if(isNeedToRemoveSheet.get(sheetName) != null && isNeedToRemoveSheet.get(sheetName)){
                    workbook.removeSheetAt(i);
                    len = workbook.getNumberOfSheets();
                    i --;
                }
            }
        }*/

        return workbook;
    }
}
