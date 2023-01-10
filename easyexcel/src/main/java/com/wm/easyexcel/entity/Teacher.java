package com.wm.easyexcel.entity;

import com.alibaba.excel.annotation.ExcelProperty;
import lombok.Data;

/**
 * @ClassName: Teacher
 * @Description: 老师
 * @Author: WM
 * @Date: 2021-07-22 21:30
 **/
@Data
public class Teacher {
    @ExcelProperty("编号")
    private String sNo;
    @ExcelProperty("老师名字")
    private String name;
    @ExcelProperty("性别")
    private String sex;
}
