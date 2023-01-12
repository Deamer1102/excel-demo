package com.wm.easyexcel.controller;

import cn.hutool.core.collection.CollUtil;
import com.wm.easyexcel.entity.Student;
import com.wm.easyexcel.handler.UserExcelHeadHandler;
import com.wm.easyexcel.service.IStudentService;
import com.wm.easyexcel.service.ITeacherService;
import com.wm.easyexcel.strategy.RowMergeStrategy;
import com.wm.easyexcel.util.CellWidthStyleHandler;
import com.wm.easyexcel.util.EasyExcelUtil;
import com.wm.easyexcel.util.ExcelHeadStyles;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import javax.annotation.Resource;
import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.*;
import java.util.stream.Collectors;

/**
 * @ClassName: ExportController
 * @Description: 导出前端控制器
 * @Author: WM
 * @Date: 2023/1/9 16:44
 */
@Slf4j
@RestController
public class ExportController {

    @Resource
    private IStudentService studentService;
    @Resource
    private ITeacherService teacherService;

    private static final String[] USER_HEAD_FIELDS = new String[]{"序号", "学号", "姓名", "性别"};

    /**
     * 导出学生数据-一个sheet
     *
     * @param response
     */
    @GetMapping("/student/export/sheet")
    public void exportSheetStudent(HttpServletResponse response) {
        List<Student> students = studentService.listStudent(10L);
        log.info("单sheet学生数据开始导出...");
        long startTimeL = System.currentTimeMillis();
        try {
            EasyExcelUtil.exportExcel(response, students, Student.class, "导出学生数据-一个sheet", "第一页");
        } catch (Exception e) {
            log.error("单sheet学生数据导出异常！", e);
        }
        log.info("单sheet学生数据导出完成：{}秒", (System.currentTimeMillis() - startTimeL) / 1000);
    }

    /**
     * 导出学生数据-多个sheet
     *
     * @param response
     */
    @GetMapping("/student/export/sheets")
    public void exportSheetsStudent(HttpServletResponse response) {
        log.info("多sheet学生数据开始导出...");
        long startTimeL = System.currentTimeMillis();
        try {
            List<List<?>> sheetList = new ArrayList<>();
            List<Student> students1 = studentService.listStudent(10L);
            List<Student> students2 = studentService.listStudent(20L);
            sheetList.add(students1);
            sheetList.add(students2);
            Map<Integer, String> clazzMap = new HashMap<>();
            clazzMap.put(0, "sheet1");
            clazzMap.put(1, "sheet2");
            EasyExcelUtil.exportExcels(response, sheetList, clazzMap, "导出学生数据-多个sheet");
        } catch (Exception e) {
            log.error("多sheet学生数据导出失败！", e);
        }
        log.info("多sheet学生数据导出完成：{}秒", (System.currentTimeMillis() - startTimeL) / 1000);
    }


    /**
     * 导出学生数据-重复行合并格式
     *
     * @param response
     */
    @GetMapping("/student/export/mergeRow")
    public void exportMergeRowStudent(HttpServletResponse response) {
        log.info("合并格式学生数据开始导出...");
        long startTimeL = System.currentTimeMillis();
        try {
            List<Student> students = studentService.listStudent(10L);
            List<String> mergeDataList = students.stream().map(Student::getSex).collect(Collectors.toList());
            EasyExcelUtil.exportWriteHandlerExcel(response, students, Student.class, "学生重复数据表格合并", "第一页", new RowMergeStrategy(mergeDataList, 3));
        } catch (Exception e) {
            log.error("合并格式学生数据导出失败！", e);
        }
        log.info("合并格式学生数据导出完成：{}秒", (System.currentTimeMillis() - startTimeL) / 1000);
    }

    /**
     * 下载导入学生数据的模板【一个自定义Handler】
     *
     * @param response
     */
    @GetMapping("/student/downTemplate001")
    public void downTemplateStudent001(HttpServletResponse response) {
        log.info("设置单元格宽度-学生数据模板开始下载...");
        long startTimeL = System.currentTimeMillis();
        try {
            List<List<String>> headList = Arrays.stream(USER_HEAD_FIELDS).map(field -> CollUtil.newArrayList(field)).
                    collect(Collectors.toList());
            CellWidthStyleHandler columnWidthHandler = new CellWidthStyleHandler(true, 15);
            EasyExcelUtil.exportExcel(response, headList, Collections.EMPTY_LIST, "学生数据导入模板-设置单元格宽度", "student", columnWidthHandler);
        } catch (Exception e) {
            log.error("设置单元格宽度-学生数据模板下载失败！", e);
        }
        log.info("设置单元格宽度-学生数据模板下载完成：{}秒", (System.currentTimeMillis() - startTimeL) / 1000);
    }

    /**
     * 下载导入学生数据的模板【多个自定义Handler】
     *
     * @param response
     */
    @GetMapping("/student/downTemplate002")
    public void downTemplateStudent002(HttpServletResponse response) {
        log.info("设置单元格宽度+字体-学生数据模板开始下载...");
        long startTimeL = System.currentTimeMillis();
        try {
            List<List<String>> headList = Arrays.stream(USER_HEAD_FIELDS).map(field -> CollUtil.newArrayList(field)).
                    collect(Collectors.toList());
            // 表头第一行的指定前3列字标红
            ExcelHeadStyles excelHeadStyle = new ExcelHeadStyles(0, 3, IndexedColors.RED1.getIndex());
            UserExcelHeadHandler headHandler = new UserExcelHeadHandler(excelHeadStyle);
            CellWidthStyleHandler columnWidthHandler = new CellWidthStyleHandler(true, 15);
            EasyExcelUtil.exportExcel(response, headList, Collections.EMPTY_LIST, "学生数据导入模板-设置单元格宽度+字体", "student", CollUtil.newArrayList(headHandler, columnWidthHandler));
        } catch (Exception e) {
            log.error("设置单元格宽度+字体-学生数据模板下载失败！", e);
        }
        log.info("设置单元格宽度+字体-学生数据模板下载完成：{}秒", (System.currentTimeMillis() - startTimeL) / 1000);
    }


    /**************以下填充模板数据的接口测试说明：可改为get请求，InputStream可使用本地文件输入流，然后变可通过浏览器下载。******************************/
    /**
     * 根据模板将集合对象填充表格-list
     *
     * @param template
     * @param response
     */
    @PostMapping("/student/template/fill-list")
    public void exportFillTemplateList(@RequestParam("template") MultipartFile template, HttpServletResponse response) {
        log.info("根据模板填充学生数据，list，文件开始下载...");
        long startTimeL = System.currentTimeMillis();
        try {
            List students = studentService.listStudent(10L);
            EasyExcelUtil.exportTemplateExcelList(template.getInputStream(), response, students, "填充学生模板数据-list", "students");
        } catch (Exception e) {
            log.error("根据模板填充学生数据，list，文件下载失败！", e);
        }
        log.info("根据模板填充学生数据，list，文件下载完成：{}秒", (System.currentTimeMillis() - startTimeL) / 1000);
    }

    /**
     * 根据模板将集合对象填充表格-单个sheet
     *
     * @param template
     * @param response
     */
    @PostMapping("/student/template/fill-sheet")
    public void exportFillTemplateSheet(@RequestParam("template") MultipartFile template, HttpServletResponse response) {
        log.info("根据模板填充学生数据，单个sheet，文件开始下载...");
        long startTimeL = System.currentTimeMillis();
        try {
            List<Student> students = studentService.listStudent(10L);
            Map<String, Object> map = new HashMap<>(3);
            map.put("startMonth", "2022-8-10");
            map.put("endMonth", "2022-8-12");
            EasyExcelUtil.exportTemplateSheet(template.getInputStream(), response, students, map, "填充学生模板数据-sheet");
        } catch (Exception e) {
            log.error("根据模板填充学生数据，单个sheet，文件下载失败！", e);
        }
        log.info("根据模板填充学生数据，单个sheet，文件下载完成：{}秒", (System.currentTimeMillis() - startTimeL) / 1000);
    }

    /**
     * 根据模板将集合对象填充表格-多个sheet
     *
     * @param template
     * @param response
     */
    @PostMapping("/student/template/fill-sheets")
    public void exportFillTemplateSheets(@RequestParam("template") MultipartFile template, HttpServletResponse response) {
        log.info("根据模板填充数据，多个sheet，文件开始下载...");
        long startTimeL = System.currentTimeMillis();
        try {
            List students = studentService.listStudent(10L);
            Map<String, Object> map = new HashMap<>(2);
            map.put("startMonth", "2022-8-10");
            map.put("endMonth", "2022-8-12");
            List teachers = teacherService.listTeacher(10L);
            Map<String, Object> map2 = new HashMap<>(1);
            map2.put("illustrate", "老师人员列表");
            EasyExcelUtil.exportTemplateSheets(template.getInputStream(), response, students, teachers, map, map2, "填充学生模板数据-sheets");
        } catch (Exception e) {
            log.error("根据模板填充数据，多个sheet，文件下载失败！", e);
        }
        log.info("根据模板填充数据，多个sheet，文件下载完成：{}秒", (System.currentTimeMillis() - startTimeL) / 1000);
    }

    /********转为get请求测试示例********/
    /**
     * get请求测试
     *
     * @param response
     */
    @GetMapping("/student/template/test")
    public void test(HttpServletResponse response) throws FileNotFoundException {
        File file = new File("C:\\Users\\Administrator\\Desktop\\test.xlsx");
        FileInputStream inputStream = new FileInputStream(file);
        log.info("根据模板填充学生数据，单个sheet，文件开始下载...");
        long startTimeL = System.currentTimeMillis();
        try {
            List<Student> students = studentService.listStudent(10L);
            Map<String, Object> map = new HashMap<>(3);
            map.put("startMonth", "2022-8-10");
            map.put("endMonth", "2022-8-12");
            EasyExcelUtil.exportTemplateSheet(inputStream, response, students, map, "填充学生模板数据-sheet");
        } catch (Exception e) {
            log.error("根据模板填充学生数据，单个sheet，文件下载失败！", e);
        }
        log.info("根据模板填充学生数据，单个sheet，文件下载完成：{}秒", (System.currentTimeMillis() - startTimeL) / 1000);
    }
}
