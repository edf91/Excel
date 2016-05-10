package org.wxd.excel.handler.inport;

import com.google.common.collect.Lists;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.List;
import java.util.Map;

/**
 * @Description: 将excel数值转实体执行器
 * @Author : wangxd
 * @Date: 2016-3-7
 * @Version 1.0
 */
public class EntityHandlerExecutor{
    private List<EntityHandler> handlers = Lists.newArrayList();

    EntityHandlerExecutor(){}

    public static EntityHandlerExecutor instance(){
        return new EntityHandlerExecutor();
    }

    public EntityHandlerExecutor register(EntityHandler entityHandler){
        this.handlers.add(entityHandler);
        return this;
    }

    public EntityHandlerExecutor handler(Workbook workbook,Map<String,List<Object>> result, Object custom){
        for (EntityHandler handler : this.handlers) handler.handlerExcelToEntity(workbook,result,custom);
        return this;
    }

}
