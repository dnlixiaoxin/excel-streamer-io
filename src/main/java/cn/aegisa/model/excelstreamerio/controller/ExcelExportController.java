package cn.aegisa.model.excelstreamerio.controller;

import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.metadata.Sheet;
import com.alibaba.excel.support.ExcelTypeEnum;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.UUID;

/**
 * @author lixiaoxin
 * @serial
 * @since 2019-04-29 11:53
 */
@Controller
@RequestMapping("excel")
public class ExcelExportController {

    @RequestMapping("/export")
    public void cooperation(HttpServletRequest request, HttpServletResponse response) {

        ServletOutputStream out = null;
        try {
            out = response.getOutputStream();
        } catch (IOException e) {
            e.printStackTrace();
        }
        ExcelWriter writer = new ExcelWriter(out, ExcelTypeEnum.XLSX, true);
        try {
            String fileName = new String(("UserInfo " + new SimpleDateFormat("yyyy-MM-dd").format(new Date()))
                    .getBytes(), StandardCharsets.UTF_8);
            Sheet sheet1 = new Sheet(1, 0);
            sheet1.setSheetName("第一个sheet");
            List<List<Object>> listString = getListString();
            writer.write1(listString, sheet1);
            response.setContentType("multipart/form-data");
            response.setCharacterEncoding("utf-8");
            response.setHeader("Content-disposition", "attachment;filename=" + fileName + ".xlsx");
            assert out != null;
            out.flush();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            writer.finish();
            try {
                assert out != null;
                out.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    private List<List<Object>> getListString() {
        List<List<Object>> data = new ArrayList<>();
        for (int i = 0; i < 255; i++) {
            for (int j = 0; j < 20; j++) {
                List<Object> row = new ArrayList<>();
                for (int k = 0; k < 4; k++) {
                    row.add(UUID.randomUUID().toString());
                }
                data.add(row);
            }
        }
       return data;
    }
}
