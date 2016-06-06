package org.wxd.excel.handler.export;

import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import org.apache.poi.ss.usermodel.*;
import org.wxd.excel.bean.CellInfo;
import org.wxd.excel.bean.ExcelContent;
import org.wxd.excel.bean.ExcelTemplate;
import org.wxd.excel.bean.ExcelTemplateParam;
import org.wxd.excel.handler.inport.ExcelHandler;
import org.wxd.excel.utils.Assert;
import org.wxd.excel.utils.ExcelUtil;

import java.math.BigDecimal;
import java.util.List;
import java.util.Map;
import java.util.concurrent.*;

/**
 * 多线程处理
 * Created by wangxd on 16/5/10.
 */
public class DefaultExportSyncHandler implements ExcelHandler {

    @SuppressWarnings({"Duplicates", "unchecked", "SpellCheckingInspection"})
    public Workbook handlerWorkbook(Workbook workbook, ExcelContent content, Object custom) {

        List<String> sheetTitles = (List<String>) custom;
        List<ExcelTemplate> excelTemplates = content.templates();
        List<ExcelTemplateParam> excelTemplateParams = content.params();

        CellStyle style = workbook.createCellStyle();
        ExcelUtil.buildDefaultStyle(style);


        Map<String, Boolean> isNeedToRemoveSheet = Maps.newHashMap();
        /*移除不必要的sheet*/
        for (int i = 0, len = workbook.getNumberOfSheets(); i < len; i++) {
            String sheetName = workbook.getSheetAt(i).getSheetName();
            if (!sheetTitles.contains(sheetName)) {
                workbook.removeSheetAt(i);
                len = workbook.getNumberOfSheets();
                i--;
                continue;
            }
            isNeedToRemoveSheet.put(sheetName, Boolean.TRUE);
        }



        /**
         * 采用多线程进行处理,一个sheet一条线程
         */
        ExecutorService exec = Executors.newCachedThreadPool();
        List<Future<Boolean>> futures = Lists.newArrayList();
        for (String sheetTitle : sheetTitles) {
            ExportHandlerRunnable exportHandlerRunnable = new ExportHandlerRunnable();
            exportHandlerRunnable.excelTemplateParams = excelTemplateParams;
            exportHandlerRunnable.workbook = workbook;
            exportHandlerRunnable.excelTemplates = excelTemplates;
            exportHandlerRunnable.sheetTitle = sheetTitle;
            exportHandlerRunnable.isNeedToRemoveSheet = isNeedToRemoveSheet;
            exportHandlerRunnable.style = style;
            futures.add(exec.submit(exportHandlerRunnable));
        }
        /**
         * 主线程判断子线程是否处理完毕
         * 主线程等待300ms再次判断
         */
        for (Future<Boolean> future : futures) {
            try {
                if(future.get() == null || !future.get()){
                    System.out.println("faile......");
                    exec.shutdown();
                }
                if (!future.isDone()) System.out.println("export not done.");
            } catch (InterruptedException e) {
                e.printStackTrace();
            } catch (ExecutionException e) {
                e.printStackTrace();
            }
        }

        return workbook;
    }

    /**
     * 处理线程类
     */
    public static class ExportHandlerRunnable implements Callable<Boolean>{

        private List<ExcelTemplate> excelTemplates;
        private List<ExcelTemplateParam> excelTemplateParams;
        private Map<String, Boolean> isNeedToRemoveSheet;
        private String sheetTitle;
        private Workbook workbook;
        private Sheet sheet;
        private Row row;
        private Cell cell;
        private CellStyle style;

        @SuppressWarnings("Duplicates")
        private void dealParam(){
            /*处理参数*/
            for (ExcelTemplateParam excelTemplateParam : excelTemplateParams) {
                if (excelTemplateParam.sheetTitle() == null || !sheetTitle.equals(excelTemplateParam.sheetTitle()) || "".equals(excelTemplateParam.sheetTitle()))
                    continue;
                sheet = workbook.getSheet(excelTemplateParam.sheetTitle());
                isNeedToRemoveSheet.remove(excelTemplateParam.sheetTitle());
                /*处理参数*/
                if(sheet == null) return ;
                for (int index = sheet.getFirstRowNum(); index <= sheet.getLastRowNum(); index++) {
                    row = sheet.getRow(index);
                    if(row == null) return;
                    for (int cellIndex = row.getFirstCellNum(); cellIndex <= row.getLastCellNum(); cellIndex++) {
                        if (cellIndex < 0) break;
                        cell = row.getCell(cellIndex);
                        if (cell == null) continue;
                        Object objValue = ExcelUtil.getCellValue(cell);
                        String value = objValue == null ? "" : objValue.toString();
                        if (!value.contains("${")) continue;
                        int valueLength = value.length();
                        int subLength;
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
        }

        @SuppressWarnings("Duplicates")
        private void dealContent(){

            Map<String, Integer> hasDealIndexMap = Maps.newHashMap();
            for (ExcelTemplate excelTemplate : excelTemplates) {
                if (excelTemplate.sheetTitle() == null || !sheetTitle.equals(excelTemplate.sheetTitle())) continue;
                Integer currentIndex = hasDealIndexMap.get(excelTemplate.sheetTitle()) == null ? excelTemplate.beginWriteRowIndex() : hasDealIndexMap.get(excelTemplate.sheetTitle());
                sheet = workbook.getSheet(excelTemplate.sheetTitle());
                row = sheet.createRow(currentIndex);
                row.setHeight((short) (20 * 18));
                for (Map.Entry<Integer, CellInfo> entry : excelTemplate.orderCellMap().entrySet()) {
                    CellInfo cellInfo = entry.getValue();
                    String cellValue = cellInfo.value() == null ? "" : cellInfo.value().toString();
                    if (cellInfo.order() == -1) continue;
                    cell = row.createCell(cellInfo.order());
                    if (cellInfo.fieldType().toString().equals("class java.math.BigDecimal")) {
                        if(cellValue.equals("")){
                            cell.setCellValue("");
                        }else{
                            cell.setCellValue(new BigDecimal(cellValue).doubleValue());
                        }
                    }else if(cellInfo.fieldType().toString().equals("class java.lang.Integer")){
                        if(cellValue.equals("")){
                            cell.setCellValue("");
                        }else{
                            cell.setCellValue(new BigDecimal(cellValue).intValue());
                        }
                    } else {
                        cell.setCellValue(cellValue);
                    }
                    if (cellInfo.styles() == null) continue;
                    cell.setCellStyle(style);
                }
                hasDealIndexMap.put(excelTemplate.sheetTitle(), ++currentIndex);
            }
        }

        @Override
        public Boolean call() throws Exception {
            Assert.notNull(excelTemplates,"excelTemplates cant be null");
            Assert.notNull(excelTemplateParams,"excelTemplateParams cant be null");
            Assert.notNull(isNeedToRemoveSheet,"isNeedToRemoveSheet cant be null");
            Assert.notNull(sheetTitle,"sheetTitle cant be null");
            Assert.notNull(workbook,"workbook cant be null");
            Assert.notNull(style,"style cant be null");
            /* 处理参数*/
            dealParam();
            /*处理内容*/
            dealContent();
            return Boolean.TRUE;
        }
    }

}
