package cn.aegisa.model.excelstreamerio.listener;

import cn.aegisa.model.excelstreamerio.vo.ExcelPropertyIndexModel;
import com.alibaba.excel.read.context.AnalysisContext;
import com.alibaba.excel.read.event.AnalysisEventListener;
import com.alibaba.fastjson.JSON;

/**
 * Using IntelliJ IDEA.
 *
 * @author XIANYINGDA at 10/12/2018 12:16 PM
 */
public class ExcelListener extends AnalysisEventListener<ExcelPropertyIndexModel> {

    @Override
    public void invoke(ExcelPropertyIndexModel object, AnalysisContext context) {
        System.out.println("当前行：" + context.getCurrentRowNum());
        System.out.println(JSON.toJSONString(object));
    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext context) {
        System.out.println("都整完啦");
    }
}
