package cn.aegisa.model.excelstreamerio.controller;

import cn.aegisa.model.excelstreamerio.utils.ExcelStreamer;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
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

    @RequestMapping("/test")
    @ResponseBody
    public String doTEstExcel(HttpServletResponse response) throws IOException {
        List<String> head = new ArrayList<>();
        head.add("name");
        head.add("age");
        head.add("add");
        head.add("phone");
        head.add("id");
        head.add("score");
        ExcelStreamer streamer = ExcelStreamer.getExcelStreamer("测试表单", head, "d:\\test.xlsx");
        for (int i = 0; i < 255; i++) {
            List<List<String>> data = new ArrayList<>();
            for (int j = 0; j < 4096; j++) {
                List<String> row = new ArrayList<>();
                for (int k = 0; k < 6; k++) {
                    row.add(UUID.randomUUID().toString());
                }
                data.add(row);
            }
            streamer.write(data);
        }
        streamer.finish();
        return "ok";
    }
}
