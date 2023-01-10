package com.wm.easyexcel.util;

import lombok.Data;

/**
 * @ClassName: ExcelHeadStyles
 * @Description: excel表头颜色定义
 * @Author: WM
 * @Date: 2022/12/22 17:24
 */
@Data
public class ExcelHeadStyles {

    /**
     * 表头横坐标 - 行
     */
    private Integer rowIndex;

    /**
     * 表头纵坐标 - 列
     */
    private Integer columnIndex;

    /**
     * 内置颜色
     */
    private Short indexColor;

    /**
     * 字体颜色
     */
    private Short fontColor;

    public ExcelHeadStyles(Integer columnIndex, Short fontColor) {
        this.columnIndex = columnIndex;
        this.fontColor = fontColor;
    }

    public ExcelHeadStyles(Integer rowIndex, Integer columnIndex, Short fontColor) {
        this.rowIndex = rowIndex;
        this.columnIndex = columnIndex;
        this.fontColor = fontColor;
    }

}
