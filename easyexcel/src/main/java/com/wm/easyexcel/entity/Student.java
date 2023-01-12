package com.wm.easyexcel.entity;

import com.alibaba.excel.annotation.ExcelProperty;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

/**
 * @ClassName: Student
 * @Description: 学生
 * @Author: WM
 * @Date: 2021-07-22 21:30
 **/
@Data
@NoArgsConstructor
@AllArgsConstructor
public class Student {
    @ExcelProperty("序号")
    private Integer num;
    @ExcelProperty("学号")
    private String sno;
    @ExcelProperty("学生名字")
    private String name;
    @ExcelProperty("性别")
    private String sex;
}
