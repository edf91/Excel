package org.wxd.excel.handler.export;

import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import org.apache.poi.ss.usermodel.*;
import org.wxd.excel.annotation.ExcelCellStyle;
import org.wxd.excel.bean.CellInfo;
import org.wxd.excel.bean.ExcelContent;
import org.wxd.excel.bean.ExcelTemplate;
import org.wxd.excel.bean.ExcelTemplateParam;
import org.wxd.excel.handler.inport.ExcelHandler;

import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.Map;

/**
 * 多线程处理
 * Created by wangxd on 16/5/10.
 */
public class DefaultExportSyncHandler implements ExcelHandler {
    /**
     * 获取但单元格值
     * @param cell
     * @return
     */
    @SuppressWarnings("Duplicates")
    public static Object getCellValue(Cell cell){
        Object cellValue = null;
        switch (cell.getCellType()) {
            case Cell.CELL_TYPE_FORMULA:
                try {
                    cellValue = cell.getStringCellValue();
                }catch (Exception e){
                    cellValue = cell.getNumericCellValue();
                }
                break;
            case Cell.CELL_TYPE_NUMERIC: // 数字,或者日期
                try{
                    cellValue = new BigDecimal(cell.getNumericCellValue());
                }catch (Exception e){
                    int format = cell.getCellStyle().getDataFormat();
                    if (DateUtil.isCellDateFormatted(cell)) {// 处理日期格式、时间格式
                        cellValue = cell.getDateCellValue();
                        if(cellValue != null) cellValue = new SimpleDateFormat("yyyy-MM-dd hh:mm:ss").format(new Date(cellValue.toString()));
                    } else if (format == 58 || format == 176 || format == 184 || format == 31) {
                        // 处理自定义日期格式：m月d日(通过判断单元格的格式id解决，id的值是58)
                        cellValue = DateUtil.getJavaDate(cell.getNumericCellValue());
                        if(cellValue != null) cellValue = new SimpleDateFormat("yyyy-MM-dd hh:mm:ss").format(new Date(cellValue.toString()));
                    } else {
                        cellValue = cell.getNumericCellValue();
                    }
                }
                break;
            case Cell.CELL_TYPE_STRING: // 字符串
                cellValue = cell.getStringCellValue();
                break;
            case Cell.CELL_TYPE_BOOLEAN: // Boolean
                cellValue = cell.getBooleanCellValue();
                break;
            case Cell.CELL_TYPE_BLANK: // 空值
                cellValue = "";
                break;
            case Cell.CELL_TYPE_ERROR: // 故障
                cellValue = "非法字符";
                break;
            default:
                cellValue = "未知类型";
                break;
        }
        return cellValue;
    }


    @SuppressWarnings("Duplicates")
    public Workbook handlerWorkbook(Workbook workbook, ExcelContent content, Object custom) {

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

        List<ExportHandlerRunnable> runnables = Lists.newArrayList();
        for (String sheetTitle : sheetTitles) {

            ExportHandlerRunnable exportHandlerRunnable = new ExportHandlerRunnable();
            exportHandlerRunnable.excelTemplateParams = excelTemplateParams;
            exportHandlerRunnable.workbook = workbook;
            exportHandlerRunnable.excelTemplates = excelTemplates;
            exportHandlerRunnable.sheetTitle = sheetTitle;
            exportHandlerRunnable.isNeedToRemoveSheet = isNeedToRemoveSheet;
            runnables.add(exportHandlerRunnable);
            new Thread(exportHandlerRunnable).start();
        }
        long start = System.currentTimeMillis();
        System.out.println("start deal:" + start);
        while(true){
            boolean isDone = true;
            for (ExportHandlerRunnable runnable : runnables) {
//                System.out.println("isFinish" + runnable.hasFinish);
                if(!runnable.hasFinish){
                    isDone = false;
                    break;
                }
            }
            if(isDone){
                break;
            }else{
                try {
                    Thread.sleep(500);
                } catch (InterruptedException e) {
                    e.printStackTrace();
                }
            }
        }
        System.out.println("finish deal:" + System.currentTimeMillis());
        System.out.println("allTime:" + (System.currentTimeMillis() - start));
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

    public static class ExportHandlerRunnable implements Runnable{

        private  List<ExcelTemplate> excelTemplates;
        List<ExcelTemplateParam> excelTemplateParams;
        Map<String,Boolean> isNeedToRemoveSheet;
        private String sheetTitle;
        private Workbook workbook;
        private boolean hasFinish = false;

        @SuppressWarnings("Duplicates")
        public void run() {
            Sheet sheet;
            Row row;
            Cell cell;
            CellStyle style = workbook.createCellStyle();

            /*处理参数*/
            for (ExcelTemplateParam excelTemplateParam : excelTemplateParams) {
                if(excelTemplateParam.sheetTitle() == null || !sheetTitle.equals(excelTemplateParam.sheetTitle()) || "".equals(excelTemplateParam.sheetTitle())) continue;
                sheet = workbook.getSheet(excelTemplateParam.sheetTitle());
                isNeedToRemoveSheet.remove(excelTemplateParam.sheetTitle());
                /*处理参数*/
                for (int index = sheet.getFirstRowNum(); index <= sheet.getLastRowNum(); index++) {
                    row = sheet.getRow(index);
                    for (int cellIndex = row.getFirstCellNum(); cellIndex <= row.getLastCellNum(); cellIndex++) {
                        if (cellIndex < 0) break;
                        cell = row.getCell(cellIndex);
                        if (cell == null) continue;
                        Object objValue = getCellValue(cell);
                        String value = objValue == null ? "" : objValue.toString();
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
            /*处理内容*/
            Map<String, Integer> hasDealIndexMap = Maps.newHashMap();
            for (ExcelTemplate excelTemplate : excelTemplates) {
                if (excelTemplate.sheetTitle() == null || !sheetTitle.equals(excelTemplate.sheetTitle())) continue;
                Integer currentIndex = hasDealIndexMap.get(excelTemplate.sheetTitle()) == null ? excelTemplate.beginWriteRowIndex() : hasDealIndexMap.get(excelTemplate.sheetTitle());
                sheet = workbook.getSheet(excelTemplate.sheetTitle());
//                System.out.println("myTitle:" + sheetTitle);

//                sheet.shiftRows(currentIndex, sheet.getLastRowNum(), 1, true, false);
                row = sheet.createRow(currentIndex);
                row.setHeight((short) (20 * 18));
                for (Map.Entry<Integer, CellInfo> entry : excelTemplate.orderCellMap().entrySet()) {
                    CellInfo cellInfo = entry.getValue();
                    String cellValue = cellInfo.value() == null ? "" : cellInfo.value().toString();
                    if (cellInfo.order() == -1) continue;
                    cell = row.createCell(cellInfo.order());
                    if(cellInfo.fieldType().toString().equals("class java.math.BigDecimal")){
                        cell.setCellValue(new BigDecimal(cellValue.equals("") ?  "0" : cellValue).doubleValue());
                    }else if(cellInfo.fieldType().toString().equals("class java.lang.Integer")){
                        cell.setCellValue(new BigDecimal(cellValue.equals("") ?  "0" : cellValue).intValue());
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
            hasFinish = true;

        }
    }

}
