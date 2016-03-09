package org.wxd.excel.bean.handler;


import org.wxd.excel.bean.ExcelRepository;

/**
 * @Description: template读取repository的句柄,用户如果想自定义读取repository可以实现该接口，并实现自己的readExcelRepository方法
 * @Author : wangxd
 * @Date: 2016-3-3
 * @Version 1.0
 */
public interface ExcelTemplateRedRepositoryHandler<T extends ExcelTemplateRedRepositoryHandler> {
    /**
     * template读取数据句柄
     * @param repository
     * @return
     */
    T readExcelRepository(ExcelRepository repository);
}
