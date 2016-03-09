package org.wxd.excel.handler;

import com.google.common.collect.Lists;
import org.apache.poi.ss.usermodel.Workbook;
import org.wxd.excel.bean.ExcelContent;

import java.util.List;

/**
 * @Description: 处理类注册，顺序执行各个处理类
 * @Author : wangxd
 * @Date: 2016-3-7
 * @Version 1.0
 */
public class ExcelHandlerExecutor {

    private List<ExcelHandler> handlers = Lists.newArrayList();


    ExcelHandlerExecutor(){}

    public static ExcelHandlerExecutor instance(){
        return new ExcelHandlerExecutor();
    }
    /**
     * 注册
     * @param handler
     * @return
     */
    public ExcelHandlerExecutor register(ExcelHandler handler){
        this.handlers.add(handler);
        return this;
    }

    /**
     * 注册
     * @param handlers
     * @return
     */
    public ExcelHandlerExecutor registerAll(List<ExcelHandler> handlers){
        this.handlers.addAll(handlers);
        return this;
    }

    public ExcelHandlerExecutor handler(Workbook workbook,ExcelContent content,Object custom){
        for (ExcelHandler handler : this.handlers) handler.handlerWorkbook(workbook,content,custom);
        return this;
    }
}
