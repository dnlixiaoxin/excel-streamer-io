package cn.aegisa.model.excelstreamerio.vo;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.metadata.BaseRowModel;
import lombok.Data;
import lombok.EqualsAndHashCode;

/**
 * Using IntelliJ IDEA.
 *
 * @author XIANYINGDA at 10/12/2018 11:13 AM
 */
@EqualsAndHashCode(callSuper = false)
@Data
public class ExcelPropertyIndexModel extends BaseRowModel {
    @ExcelProperty(value = "姓名", index = 0)
    private String name;

    @ExcelProperty(value = "年龄", index = 1)
    private String age;

    @ExcelProperty(value = "邮箱", index = 2)
    private String email;

    @ExcelProperty(value = "地址", index = 3)
    private String address;
}
