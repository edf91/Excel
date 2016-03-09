package org.wxd.excel.utils;

import java.util.Comparator;

/**
 * @Description: 数值排序，越小越前
 * @Copyright: Copyright 2012 ShenZhen DSE Corporation
 * @Company: 深圳市东深电子股份有限公司
 * @Author : wangxd
 * @Date: 2016-3-2
 * @Version 1.0
 */
public class NumberSortComparator implements Comparator<Integer> {
    public boolean isDesc;

    public NumberSortComparator(boolean isDesc) {
        this.isDesc = isDesc;
    }

    @Override
    public int compare(Integer o1, Integer o2) {
        return isDesc ? o2 - o1 : o1 - o2;
    }
}
