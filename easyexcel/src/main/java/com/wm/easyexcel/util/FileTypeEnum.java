package com.wm.easyexcel.util;

import lombok.Getter;

/**
 * @ClassName: FileTypeEnum
 * @Description: 文件类型枚举
 * @Author: WM
 * @Date: 2022/12/21 18:19
 */
@Getter
public enum FileTypeEnum {
    /**
     * 模板格式
     */
    TEMPLATE_SUFFIX("xlsx", ".xlsx"),
    TEMPLATE_SUFFIX_XLS("xls", ".xls"),
    TEMPLATE_SUFFIX_DOCX("docx", ".docx");
    private final String code;
    private final String desc;

    FileTypeEnum(String code, String desc) {
        this.code = code;
        this.desc = desc;
    }

    /**
     * 通过code获取msg
     *
     * @param code 枚举值
     * @return
     */
    public static String getMsgByCode(String code) {
        if (code == null) {
            return null;
        }
        FileTypeEnum enumList = getByCode(code);
        if (enumList == null) {
            return null;
        }
        return enumList.getDesc();
    }

    public static String getCode(FileTypeEnum enumList) {
        if (enumList == null) {
            return null;
        }
        return enumList.getCode();
    }

    public static FileTypeEnum getByCode(String code) {
        for (FileTypeEnum enumList : values()) {
            if (enumList.getCode().equals(code)) {
                return enumList;
            }
        }
        return null;
    }
}
