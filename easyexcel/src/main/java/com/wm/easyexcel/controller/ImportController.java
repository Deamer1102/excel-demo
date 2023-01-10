package com.wm.easyexcel.controller;

import cn.hutool.core.collection.CollUtil;
import cn.hutool.core.lang.Assert;
import com.wm.easyexcel.entity.Student;
import com.wm.easyexcel.entity.Teacher;
import com.wm.easyexcel.listern.ImportStudentListener;
import com.wm.easyexcel.util.EasyExcelUtil;
import lombok.extern.slf4j.Slf4j;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import javax.annotation.Resource;
import java.util.ArrayList;
import java.util.List;

/**
 * @ClassName: ImportController
 * @Description: 导入前端控制器
 * @Author: WM
 * @Date: 2023/1/9 16:44
 */
@Slf4j
@RestController
public class ImportController {

    @Resource
    private ImportStudentListener importStudentListener;

    /**
     * 导入一个sheet学生数据
     *
     * @param file
     */
    @PostMapping("/student/import/sheet")
    public boolean importSheetStudent(@RequestParam("file") MultipartFile file) {
        EasyExcelUtil.checkFile(file);
        log.info("一个sheet学生数据开始导入...");
        long startTime = System.currentTimeMillis();
        List<Student> dataList = null;
        try {
            dataList = EasyExcelUtil.importExcel(file.getInputStream(), null, null, importStudentListener);
            // 数据录入
            Assert.isFalse(CollUtil.isEmpty(dataList), "未解析到导入的数据！");
        } catch (Exception e) {
            log.error("一个sheet学生数据导入失败！", e);
        }
        log.info("一个sheet学生数据导入成功，共：{}条，消耗总时间：{}秒", dataList.size(), ((System.currentTimeMillis() - startTime) / 1000));
        return true;
    }

    /**
     * 导入多个sheet学生数据
     *
     * @param file
     */
    @PostMapping("/student/import/sheets")
    public boolean importSheets(@RequestParam("file") MultipartFile file) {
        EasyExcelUtil.checkFile(file);
        log.info("多个sheet数据开始导入...");
        long startTime = System.currentTimeMillis();
        List<Student> dataList = null;
        try {
            List sheetObjList = new ArrayList(2);
            sheetObjList.add(Student.class);
            sheetObjList.add(Teacher.class);
            dataList = EasyExcelUtil.importExcels(file.getInputStream(), 2, sheetObjList, importStudentListener);
            // 数据录入
            Assert.isFalse(CollUtil.isEmpty(dataList), "未解析到导入的数据！");
        } catch (Exception e) {
            log.error("多个sheet数据导入失败！", e);
        }
        log.info("多个sheet数据导入成功，共：{}条，消耗总时间：{}秒", dataList.size(), ((System.currentTimeMillis() - startTime) / 1000));
        return true;
    }
}
