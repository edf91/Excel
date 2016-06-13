package org.wxd.excel.handler.inport;

import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import org.apache.poi.ss.usermodel.*;
import org.wxd.excel.annotation.ExcelCell;
import org.wxd.excel.annotation.ExcelParam;
import org.wxd.excel.annotation.ExcelSheet;
import org.wxd.excel.exception.ExcelException;

import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.Map;

/**
 * @Description: 根据加进来的类的注解创建对象
 * @Author : wangxd
 * @Date: 2016-3-7
 * @Version 1.0
 */
public class DefaultEntityHandler implements EntityHandler{
    private List<Class> entityClasses = Lists.newArrayList();

    public DefaultEntityHandler register(Class clazz){
        this.entityClasses.add(clazz);
        return this;
    }

    public DefaultEntityHandler registerAll(List<Class> clazzes){
        this.entityClasses.addAll(clazzes);
        return this;
    }

    public Map<String, List<Object>> handlerExcelToEntity(Workbook workbook, Map<String,List<Object>> result, Object custom) {
        try {
            Map<String,Class> sheetTitleEntityClassMap = Maps.newConcurrentMap();/*存放标题对应的对象*/
            /*获取注册进来且有注解的字节码*/
            for (Class entityClass : entityClasses) {
                if(!entityClass.isAnnotationPresent(ExcelSheet.class)) continue;
                ExcelSheet excelSheet = (ExcelSheet) entityClass.getAnnotation(ExcelSheet.class);
                sheetTitleEntityClassMap.put(excelSheet.name(),entityClass);
            }
            Sheet sheet;
            for (Map.Entry<String, Class> sheetTitleClassEntry : sheetTitleEntityClassMap.entrySet()) {
                String sheetTitle = sheetTitleClassEntry.getKey();
                Class clazz = sheetTitleClassEntry.getValue();
                ExcelSheet excelSheet = (ExcelSheet) clazz.getAnnotation(ExcelSheet.class);
                sheet = workbook.getSheet(sheetTitle);
                if(sheet == null) continue;
                List<Object> entities = Lists.newArrayList();
                /*处理参数部分*/
                /*处理实体*/
                for(int i = excelSheet.beginReadRowIndex(); i <= sheet.getLastRowNum(); i++){
                    Object entity = clazz.newInstance();
                    /*处理父类属性*/
                    reflectSetFieldValue(clazz.getSuperclass().getDeclaredFields(),sheet,entity,i);
                    /*处理本身属性*/
                    reflectSetFieldValue(clazz.getDeclaredFields(),sheet,entity,i);
                    entities.add(entity);
                }
                result.put(sheetTitle,entities);
            }
            return result;
        }catch (Exception e){throw new ExcelException(e.getMessage(),e);}
    }

    private void reflectSetFieldValue(Field[] fields, Sheet sheet, Object entity, int index) throws IllegalAccessException {
        Cell cell = null;
        Row row = sheet.getRow(index);
        for (Field field : fields) {
            field.setAccessible(true);
            if(!field.isAnnotationPresent(ExcelCell.class) && !field.isAnnotationPresent(ExcelParam.class)) continue;
            if(field.isAnnotationPresent(ExcelCell.class)) cell = row.getCell(field.getAnnotation(ExcelCell.class).order());
            else if(field.isAnnotationPresent(ExcelParam.class)){
                String[] readIndex = field.getAnnotation(ExcelParam.class).readIndex().split(":");
                if(readIndex.length != 2) continue;
                cell = sheet.getRow(Integer.parseInt(readIndex[0])).getCell(Integer.parseInt(readIndex[1]));
            }
            Object cellValue;
            if(cell == null) continue;
            switch (cell.getCellType()) {
                case Cell.CELL_TYPE_NUMERIC: // 数字,或者日期
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
                    break;
                case Cell.CELL_TYPE_STRING: // 字符串
                    cellValue = cell.getStringCellValue();
                    break;
                case Cell.CELL_TYPE_BOOLEAN: // Boolean
                    cellValue = cell.getBooleanCellValue();
                    break;
                case Cell.CELL_TYPE_FORMULA: // 公式
                    cellValue = cell.getCellFormula();
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
            String value = cellValue.toString();
            if(field.getGenericType().toString().equals("class java.lang.Integer")){
                field.set(entity, value.equals("") ? null : new BigDecimal(value).intValue());
            }else if(field.getGenericType().toString().equals("class java.lang.String")){
                field.set(entity,value);
            }else if(field.getGenericType().toString().equals("class java.math.BigDecimal")){
                field.set(entity,value.equals("") ? null : new BigDecimal(value));
            }
        }
    }
}
