package org.wxd.excel.model;

import org.wxd.excel.annotation.ExcelCell;
import org.wxd.excel.annotation.ExcelCellStyle;
import org.wxd.excel.annotation.ExcelSheet;

/**
 * @Description:
 * @Copyright: Copyright 2012 ShenZhen DSE Corporation
 * @Company: 深圳市东深电子股份有限公司
 * @Author : wangxd
 * @Date: 2016-5-10
 * @Version 1.0
 */
@ExcelSheet(name = "tea",beginWriteRowIndex = 1)
public class Teacher {
    @ExcelCell(order = 0,style = {ExcelCellStyle.ALIGN_CENTER,ExcelCellStyle.BORDER_ALL})
    private Integer tNo;
    @ExcelCell(order = 1,style = {ExcelCellStyle.ALIGN_CENTER,ExcelCellStyle.BORDER_ALL})
    private String name;
    @ExcelCell(order = 2,style = {ExcelCellStyle.ALIGN_CENTER,ExcelCellStyle.BORDER_ALL})
    private String sex;

    public Teacher() {
    }

    public Teacher(Integer tNo, String name, String sex) {
        this.tNo = tNo;
        this.name = name;
        this.sex = sex;
    }

    public Integer gettNo() {
        return tNo;
    }

    public void settNo(Integer tNo) {
        this.tNo = tNo;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getSex() {
        return sex;
    }

    public void setSex(String sex) {
        this.sex = sex;
    }
}
