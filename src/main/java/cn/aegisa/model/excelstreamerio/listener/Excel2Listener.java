package cn.aegisa.model.excelstreamerio.listener;

import cn.aegisa.model.excelstreamerio.vo.ExcelPropertyIndexModel;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.fastjson.JSON;

/**
 * @author lixiaoxin
 * @serial
 * @since 2019-04-28 19:38
 */
public class Excel2Listener extends AnalysisEventListener<ExcelPropertyIndexModel> {

    @Override
    public void invoke(ExcelPropertyIndexModel excelPropertyIndexModel, AnalysisContext analysisContext) {
        System.out.println("当前执行行为：" + analysisContext.getCurrentRowNum());
        System.out.println("执行的数据为：" + JSON.toJSONString(excelPropertyIndexModel));
    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext analysisContext) {
        System.out.println("全都执行完了");
    }
}
