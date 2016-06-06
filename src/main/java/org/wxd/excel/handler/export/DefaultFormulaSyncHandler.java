package org.wxd.excel.handler.export;

import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.wxd.excel.annotation.ExcelFormula;
import org.wxd.excel.bean.ExcelContent;
import org.wxd.excel.bean.ExcelTemplateFormula;
import org.wxd.excel.exception.ExcelException;
import org.wxd.excel.handler.inport.ExcelHandler;

import java.util.List;
import java.util.Map;
import java.util.concurrent.*;

/**
 * @Description: 异步处理公式
 * @Copyright: Copyright 2012 ShenZhen DSE Corporation
 * @Company: 深圳市东深电子股份有限公司
 * @Author : wangxd
 * @Date: 2016-5-11
 * @Version 1.0
 */
public class DefaultFormulaSyncHandler implements ExcelHandler {
    @Override
    public Workbook handlerWorkbook(Workbook workbook, ExcelContent content, Object custom) {
        /**
         * 获取需要处理的sheet
         */
        Map<String,List<ExcelFormula>> formulaMap = Maps.newHashMap();
        for (ExcelTemplateFormula formula : content.formulas()) {
            if(formulaMap.get(formula.sheetTitle()) != null) continue;
            formulaMap.put(formula.sheetTitle(),formula.formulas());
        }

        ExecutorService exec = Executors.newCachedThreadPool();
        List<Future<Boolean>> futures = Lists.newArrayList();


        for (Map.Entry<String, List<ExcelFormula>> titleFormula : formulaMap.entrySet()) {
            String title = titleFormula.getKey();
            ExportHandlerRunnable exportHandlerRunnable = new ExportHandlerRunnable();
            exportHandlerRunnable.title = title;
            exportHandlerRunnable.workbook = workbook;
            exportHandlerRunnable.titleFormula = titleFormula;
            futures.add(exec.submit(exportHandlerRunnable));
        }

        /**
         * 主线程判断子线程是否处理完毕
         * 主线程等待300ms再次判断
         */
        for (Future<Boolean> future : futures) {
            try {
                if(future.isDone()){
                    System.out.println("formula is done.");
                    exec.shutdown();
                }
                if(future.get() == null || !future.get()){
                    System.out.println("formula is faile.");
                    exec.shutdown();
                }
            }catch (Exception e){
                throw new ExcelException(e.getMessage(),e);
            }
        }

        return workbook;
    }
    public static class ExportHandlerRunnable implements Callable<Boolean> {
        Map.Entry<String, List<ExcelFormula>> titleFormula;
        String title;
        Workbook workbook;

        private void dealFormula(){
            Sheet sheet;
            Row row;
            Cell cell;

            sheet = workbook.getSheet(title);
            List<ExcelFormula> formulas = titleFormula.getValue();
            for (ExcelFormula formula : formulas) {
                String index = formula.index();
                Integer rowIndex = Integer.parseInt(index.split(":")[0]);
                Integer cellIndex = Integer.parseInt(index.split(":")[1]);

                row = sheet.getRow(rowIndex);
                cell = row.getCell(cellIndex);

                /*设置公式*/
                String calc = formula.formula();
                if(calc.contains("*")){
                    calc = calc.replace("*","9999");
                    cell.setCellFormula(calc);
                    continue;
                }

                if(formula.isApplyColum()){
                    if(calc.contains(",")){
                        //SUM(K${6},N${6},O${6},P${6})
                        String numStr = calc.replaceAll("\\D+", "");
                        Integer calcRowIndex;
                        calcRowIndex = Integer.parseInt(numStr.substring(0, 1));
                        String charNum = "${" + calcRowIndex +"}";
                        for(int start = rowIndex,len = sheet.getLastRowNum() - 1; start <= len; start++,calcRowIndex ++){
                            row = sheet.getRow(start);
                            cell = row.getCell(cellIndex);
                            cell.setCellFormula(calc.replace(charNum, calcRowIndex + ""));
                        }
                        continue;
                    }

                    //SUM(表8!M${6}:O%{6})
                    if(calc.contains(":") && calc.contains("!")){
                        String newCalc = calc.split("!")[1];
                        String otherSheet = calc.split("!")[0];
                        if(workbook.getSheet(otherSheet.substring(otherSheet.length() - 2,otherSheet.length())) == null) continue;
                        Integer calcRowIndex = Integer.parseInt(newCalc.split(":")[0].replaceAll("\\D+", ""));
                        Integer calcCellIndex = Integer.parseInt(newCalc.split(":")[1].replaceAll("\\D+", ""));
                        String replaceRowIndex = "${" + calcRowIndex + "}";
                        String replaceCellIndex = "%{" + calcCellIndex + "}";
                        for(int start = rowIndex,len = sheet.getLastRowNum() - 1; start <= len; start++,calcCellIndex ++,calcRowIndex ++){
                            row = sheet.getRow(start);
                            cell = row.getCell(cellIndex);
                            String temp = calc.replace(replaceRowIndex,calcRowIndex + "");
                            temp = temp.replace(replaceCellIndex,calcCellIndex + "");
                            cell.setCellFormula(temp);
                        }
                        continue;
                    }
                    //SUM(M${6}:O%{6})
                    if(calc.contains(":")){
                        Integer calcRowIndex = Integer.parseInt(calc.split(":")[0].replaceAll("\\D+", ""));
                        Integer calcCellIndex = Integer.parseInt(calc.split(":")[1].replaceAll("\\D+", ""));
                        String replaceRowIndex = "${" + calcRowIndex + "}";
                        String replaceCellIndex = "%{" + calcCellIndex + "}";
                        for(int start = rowIndex,len = sheet.getLastRowNum() - 1; start <= len; start++,calcCellIndex ++,calcRowIndex ++){
                            row = sheet.getRow(start);
                            cell = row.getCell(cellIndex);
                            String temp = calc.replace(replaceRowIndex,calcRowIndex + "");
                            temp = temp.replace(replaceCellIndex,calcCellIndex + "");
                            cell.setCellFormula(temp);
                        }
                        continue;
                    }
                    // 表8!M${6}SUM(表1
                    if(calc.contains("!")){
                        String info = calc.split("!")[1];
                        String otherSheet = calc.split("!")[0];
                        if(workbook.getSheet(otherSheet.substring(otherSheet.length() - 2,otherSheet.length())) == null) continue;
                        Integer calcCellIndex = Integer.parseInt(info.replaceAll("\\D+", ""));
                        String replaceCellIndex = "${" + calcCellIndex + "}";
                        for(int start = rowIndex,len = sheet.getLastRowNum() - 1; start <= len; start++,calcCellIndex ++){
                            row = sheet.getRow(start);
                            cell = row.getCell(cellIndex);
                            String temp = calc.replace(replaceCellIndex,calcCellIndex + "");
                            cell.setCellFormula(temp);
                        }
                    }
                }
            }
        }

        @Override
        public Boolean call() throws Exception {
           dealFormula();
            return Boolean.TRUE;
        }
    }

}
