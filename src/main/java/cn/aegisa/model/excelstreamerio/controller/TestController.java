package cn.aegisa.model.excelstreamerio.controller;

import cn.aegisa.model.excelstreamerio.vo.ExcelPropertyIndexModel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.metadata.Sheet;
import com.alibaba.excel.support.ExcelTypeEnum;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;

import javax.servlet.http.HttpServletResponse;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.UUID;

/**
 * Using IntelliJ IDEA.
 *
 * @author XIANYINGDA at 2018/12/3 17:52
 */
@Controller
@Slf4j
public class TestController {

    @RequestMapping("/excel")
    @ResponseBody
    public String doTEstExcel(HttpServletResponse response) throws IOException {
        OutputStream out = new FileOutputStream("d:/100.xlsx");
        try {
            ExcelWriter writer = new ExcelWriter(out, ExcelTypeEnum.XLSX);
            //写第一个sheet, sheet1  数据全是List<String> 无模型映射关系
            Sheet sheet1 = new Sheet(1, 0, ExcelPropertyIndexModel.class);
            sheet1.setSheetName("第一个sheet");
            for (int i = 0; i < 1000; i++) {
                List<ExcelPropertyIndexModel> orderMainList = new ArrayList<>(1000);
                for (int j = 0; j < 1000; j++) {
                    ExcelPropertyIndexModel orderMain = new ExcelPropertyIndexModel();
                    orderMain.setName(UUID.randomUUID().toString());
                    orderMain.setAddress(UUID.randomUUID().toString());
                    orderMain.setAge(String.valueOf((double) (j + 10)));
                    orderMain.setEmail(UUID.randomUUID().toString());
                    orderMain.setHeigh(UUID.randomUUID().toString());
                    orderMain.setLast(String.valueOf(i * 1000 + j));
                    orderMainList.add(orderMain);
                }
                System.out.println(i);
                writer.write(orderMainList, sheet1);
            }
            writer.finish();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                out.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return "ok";
    }
}
