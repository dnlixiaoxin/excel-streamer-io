package com.lxx.test;

import cn.aegisa.model.excelstreamerio.utils.ExcelStreamer;
import com.alibaba.excel.ExcelReader;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import org.junit.Test;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.UUID;

/**
 * @author lixiaoxin
 * @serial
 * @since 2019-04-28 19:40
 */
public class ExcelTest {

    @Test
    public void testExportExcel() throws IOException {
        List<String> head = new ArrayList<>();
        head.add("name");
        head.add("age");
        head.add("email");
        head.add("address");
        ExcelStreamer streamer = ExcelStreamer.getExcelStreamer("测试表单", head, "/Users/lixiaoxin/Desktop/lixiaoxin.xls");
        for (int i = 0; i < 255; i++) {
            List<List<String>> data = new ArrayList<>();
            for (int j = 0; j < 20; j++) {
                List<String> row = new ArrayList<>();
                for (int k = 0; k < 4; k++) {
                    row.add(UUID.randomUUID().toString());
                }
                data.add(row);
            }
            streamer.write(data);
        }
        streamer.finish();

    }

    @Test
    public void testExcelUtil() throws IOException {
        //  写入文件
        InputStream inputStream = new FileInputStream("/Users/lixiaoxin/Desktop/222.xlsx");

        try {
            ExcelReader reader = new ExcelReader(inputStream, null,
                    new AnalysisEventListener<List<String>>() {
                        @Override
                        public void invoke(List<String> object, AnalysisContext context) {
                            System.out.println(
                                    "当前sheet:" + context.getCurrentSheet().getSheetNo() + " 当前行：" + context.getCurrentRowNum()
                                            + " data:" + object);
                            System.out.println(1111111);
                        }
                        @Override
                        public void doAfterAllAnalysed(AnalysisContext context) {

                        }
                    });

            reader.read();
        } catch (Exception e) {
            e.printStackTrace();

        } finally {
            try {
                inputStream.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    @Test
    public void testExportExcel1(){

    }

}
