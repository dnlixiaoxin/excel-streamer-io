package cn.aegisa.model.excelstreamerio;

import cn.aegisa.model.excelstreamerio.listener.ExcelListener;
import cn.aegisa.model.excelstreamerio.utils.ExcelStreamer;
import cn.aegisa.model.excelstreamerio.vo.ExcelPropertyIndexModel;
import com.alibaba.excel.ExcelReader;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.metadata.Sheet;
import com.alibaba.excel.read.event.AnalysisEventListener;
import com.alibaba.excel.support.ExcelTypeEnum;
import org.junit.Test;

import java.io.*;
import java.util.ArrayList;
import java.util.LinkedList;
import java.util.List;
import java.util.UUID;

/**
 * Using IntelliJ IDEA.
 *
 * @author XIANYINGDA at 10/12/2018 10:32 AM
 */
@SuppressWarnings("ALL")
public class TestExcelStreamer {

    @Test
    public void test05() throws Exception {
        List<String> list = new LinkedList<>();
        list.add("name");
        list.add("age");
        list.add("score");
        ExcelStreamer streamer = ExcelStreamer.getExcelStreamer(list, "d:/ff.xlsx");
        streamer.write();
        streamer.finish();
    }


    /**
     * 测试读
     */
    @Test
    public void test01() throws FileNotFoundException {
        InputStream inputStream = new FileInputStream("d:/100.xlsx");
        try {
            // 解析每行结果在listener中处理
            AnalysisEventListener listener = new ExcelListener();
            ExcelReader excelReader = new ExcelReader(inputStream, ExcelTypeEnum.XLSX, null, listener);
            excelReader.read(new Sheet(1, 1, ExcelPropertyIndexModel.class));
        } catch (Exception e) {

        } finally {
            try {
                inputStream.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    /**
     * 测试写
     *
     * @throws FileNotFoundException
     */
    @Test
    public void test02() throws FileNotFoundException {
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
    }
}
