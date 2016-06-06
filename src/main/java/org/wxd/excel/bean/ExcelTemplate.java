package org.wxd.excel.bean;

import com.google.common.collect.Maps;
import org.wxd.excel.annotation.ExcelCell;
import org.wxd.excel.bean.handler.ExcelTemplateRedRepositoryHandler;
import org.wxd.excel.exception.ExcelException;
import org.wxd.excel.utils.NumberSortComparator;
import org.wxd.excel.utils.Assert;

import java.lang.reflect.Field;
import java.util.Arrays;
import java.util.List;
import java.util.Map;

/**
 * @Description: 将对象根据注解，将其值存到改对象中
 * @Author : wangxd
 * @Date: 2016-3-2
 * @Version 1.0
 */
public class ExcelTemplate {
    private String sheetTitle;/*sheet名称*/
    private Integer beginWriteRowIndex;/*开始输出的行下标*/
    private Integer beginReadRowIndex;/*开始读取的行下标*/
    private Map<Integer, CellInfo> orderCellMap = Maps.newTreeMap(new NumberSortComparator(false));/*序号对应的单元格*/

    ExcelTemplate(Builder builder) {
        this.sheetTitle = builder.sheetTitle;
        this.beginWriteRowIndex = builder.beginWriteRowIndex;
        this.beginReadRowIndex = builder.beginReadRowIndex;
        this.orderCellMap = builder.orderCellMap;
    }

    public static ExcelTemplate.Builder newBuilder() {
        return new Builder();
    }

    public static class Builder implements ExcelTemplateRedRepositoryHandler<Builder> {
        String sheetTitle;
        Integer beginWriteRowIndex;
        private Integer beginReadRowIndex;
        Map<Integer, CellInfo> orderCellMap = Maps.newTreeMap(new NumberSortComparator(false));/*序号对应的值*/

        public ExcelTemplate build() {
            return new ExcelTemplate(this);
        }

        /**
         * 读取ExcelRepository对象，拿取里面的数据
         * @param repository
         * @return
         */
        @Override
        public Builder readExcelRepository(ExcelRepository repository) {
            Assert.notNull(repository, "repository can not be null.");
            try {
                if(repository.excelSheet() == null) return this;
                this.sheetTitle = repository.excelSheet().name();
                this.beginWriteRowIndex = repository.excelSheet().beginWriteRowIndex();
                this.beginReadRowIndex = repository.excelSheet().beginReadRowIndex();
                /*读取本身属性*/
                this.reflectFieldToMap(repository.fields(),repository.target());
                /*父类属性*/
                this.reflectFieldToMap(Arrays.asList(repository.targetClass().getSuperclass().getDeclaredFields()),repository.target());
                return this;
            } catch (Exception e) {
                throw new ExcelException(e.getMessage(), e);
            }
        }

        /**
         * 反射读取属性
         * @param fields
         * @param target
         */
        private void reflectFieldToMap(List<Field> fields,Object target){
            try{
                for (Field field : fields) {
                    field.setAccessible(true);
                    if (!field.isAnnotationPresent(ExcelCell.class)) continue;
                    ExcelCell excelCell = field.getAnnotation(ExcelCell.class);
                    this.orderCellMap.put(
                            excelCell.order(),
                            CellInfo.newBuilder()
                                    .order(excelCell.order()).styles(excelCell.style()).value(field.get(target))
                                    .fieldType(field.getType())
                                    .build()
                    );
                }
            }catch (Exception e){
                throw new ExcelException(e.getMessage(),e);
            }
        }


        public Builder sheetTitle(String sheetTitle) {
            this.sheetTitle = sheetTitle;
            return this;
        }

        public Builder beginWriteRowIndex(Integer beginWriteRowIndex) {
            this.beginWriteRowIndex = beginWriteRowIndex;
            return this;
        }

        public Builder beginReadRowIndex(Integer beginReadRowIndex) {
            this.beginReadRowIndex = beginReadRowIndex;
            return this;
        }
    }


    public String sheetTitle() {
        return sheetTitle;
    }

    public Integer beginReadRowIndex() {
        return beginReadRowIndex;
    }

    public Integer beginWriteRowIndex() {
        return beginWriteRowIndex;
    }

    public Map<Integer, CellInfo> orderCellMap() {
        return orderCellMap;
    }

    @Override
    public String toString() {
        return "ExcelTemplate{" +
                "sheetTitle='" + sheetTitle + '\'' +
                ", beginWriteRowIndex=" + beginWriteRowIndex +
                ", beginReadRowIndex=" + beginReadRowIndex +
                ", orderCellMap=" + orderCellMap +
                '}';
    }

}
