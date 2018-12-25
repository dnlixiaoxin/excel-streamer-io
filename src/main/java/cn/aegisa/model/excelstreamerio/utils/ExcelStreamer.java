package cn.aegisa.model.excelstreamerio.utils;

import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.metadata.Sheet;
import com.alibaba.excel.support.ExcelTypeEnum;

import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.Collections;
import java.util.List;
import java.util.stream.Collectors;

/**
 * Using IntelliJ IDEA.
 *
 * @author XIANYINGDA at 2018/12/25 16:45
 */
public class ExcelStreamer {

    private ExcelWriter writer;
    private Sheet sheet;
    private boolean toFile = false;

    private ExcelStreamer(ExcelWriter writer, Sheet sheet, boolean toFile) {
        this.writer = writer;
        this.sheet = sheet;
        this.toFile = toFile;
    }

    public void write() {
        writer.write(Collections.EMPTY_LIST, sheet);
    }

    public void finish() {
        writer.finish();
    }

    public static ExcelStreamer getExcelStreamer(List<String> head, OutputStream out) throws Exception {
        ExcelWriter writer = new ExcelWriter(out, ExcelTypeEnum.XLSX);
        Sheet sheet = new Sheet(0, 1);
        List<List<String>> headList = head.stream().map(Collections::singletonList).collect(Collectors.toList());
        sheet.setHead(headList);
        return new ExcelStreamer(writer, sheet, true);
    }

    public static ExcelStreamer getExcelStreamer(List<String> head, String filePath) throws Exception {
        OutputStream out = new FileOutputStream(filePath);
        return getExcelStreamer(head, out);
    }

}
