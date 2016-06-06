package org.wxd.excel.utils;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Workbook;
import org.wxd.excel.exception.ExcelException;
import org.wxd.excel.handler.inport.ExcelHandler;

import java.io.File;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.Date;
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

    /**
     * 处理大数据导出接口，采用SXSSFWorkbook处理
     * @param src
     * @param sheetTiltles
     * @param excelTemplateFile
     * @param flushNum 缓存单元格数量
     * @param handlers
     * @return
     */
    public static Workbook buildWorkbookFromEntityOfFileWithSXSSFWorkbook(Object src,List<String> sheetTiltles, File excelTemplateFile,int flushNum,ExcelHandler... handlers) {
        try {
            Workbook workbook = flushNum == 0 ?
                    ExportFatory.buildWorkbookOfTemplateWithSXSSFWorkbook(excelTemplateFile)
                    :
                    ExportFatory.buildWorkbookOfTemplateWithSXSSFWorkbook(excelTemplateFile, flushNum);
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

    /**
     * 处理大数据导出接口，采用SXSSFWorkbook处理，采用默认处理缓存数
     * @param src
     * @param sheetTiltles
     * @param excelTemplateFile
     * @param handlers
     * @return
     */
    public static Workbook buildWorkbookFromEntityOfFileWithSXSSFWorkbook(Object src,List<String> sheetTiltles, File excelTemplateFile,ExcelHandler... handlers) {
        try {
            return ExcelUtil.buildWorkbookFromEntityOfFileWithSXSSFWorkbook(src,sheetTiltles,excelTemplateFile,0,handlers);
        } catch (Exception e) {
            e.printStackTrace();
            throw new ExcelException(e.getMessage(),e);
        }
    }

    /**
     * TODO 有待优化
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

    /**
     * 构建默认样式，边框+居中
     * @param style
     */
    public static void buildDefaultStyle(CellStyle style) {
        style.setBorderBottom(CellStyle.BORDER_THIN); //下边框
        style.setBorderLeft(CellStyle.BORDER_THIN);//左边框
        style.setBorderTop(CellStyle.BORDER_THIN);//上边框
        style.setBorderRight(CellStyle.BORDER_THIN);//右边框
        style.setAlignment(CellStyle.ALIGN_CENTER);
    }
}
