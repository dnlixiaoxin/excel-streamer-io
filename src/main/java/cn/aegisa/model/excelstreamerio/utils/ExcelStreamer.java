package cn.aegisa.model.excelstreamerio.utils;

import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.metadata.Sheet;
import com.alibaba.excel.support.ExcelTypeEnum;
import lombok.extern.slf4j.Slf4j;
import org.springframework.util.StringUtils;

import java.io.FileNotFoundException;
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
@Slf4j
public class ExcelStreamer {

    private ExcelWriter writer;
    private Sheet sheet;
    private boolean toFile = false;
    private int size = 0;
    private boolean isDone = false;
    private OutputStream out;

    private ExcelStreamer(ExcelWriter writer, Sheet sheet, boolean toFile, OutputStream out) {
        this.writer = writer;
        this.sheet = sheet;
        this.toFile = toFile;
        if (toFile) {
            this.out = out;
        }
    }

    public void write(List<List<String>> dataMatrix) {
        if (isDone) {
            throw new RuntimeException("This excel has been done and can not be appended any more.");
        }
        if (dataMatrix != null && dataMatrix.size() > 0) {
            writer.write0(dataMatrix, sheet);
            this.size += dataMatrix.size();
            log.info("Write {} rows, total rows: {}.", dataMatrix.size(), this.size);
        }
    }

    public void finish() {
        log.info("Committing {} rows...", size);
        writer.finish();
        isDone = true;
        log.info("Commit {} rows to excel.", size);
        if (toFile) {
            try {
                out.close();
            } catch (Exception e) {
                log.error("Exception while closing the FileOutputStream", e);
            }
        }
    }

    private static ExcelStreamer getExcelStreamer0(List<String> head, OutputStream out, boolean toFile, String sheetName) {
        ExcelWriter writer = new ExcelWriter(out, ExcelTypeEnum.XLSX);
        Sheet sheet = new Sheet(0, 1);
        if (!StringUtils.isEmpty(sheetName)) {
            sheet.setSheetName(sheetName);
        }
        List<List<String>> headList = head.stream().map(Collections::singletonList).collect(Collectors.toList());
        sheet.setHead(headList);
        return new ExcelStreamer(writer, sheet, toFile, out);
    }

    public static ExcelStreamer getExcelStreamer(List<String> head, String filePath) throws Exception {
        OutputStream out = new FileOutputStream(filePath);
        return getExcelStreamer0(head, out, true, null);
    }

    public static ExcelStreamer getExcelStreamer(String sheetName, List<String> head, String filePath) {
        OutputStream out = null;
        try {
            out = new FileOutputStream(filePath);
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        }
        return getExcelStreamer0(head, out, true, sheetName);
    }

    public static ExcelStreamer getExcelStreamer(List<String> head, OutputStream out) throws Exception {
        return getExcelStreamer0(head, out, false, null);
    }

    public static ExcelStreamer getExcelStreamer(String sheetName, List<String> head, OutputStream out) {
        return getExcelStreamer0(head, out, false, sheetName);
    }

}
