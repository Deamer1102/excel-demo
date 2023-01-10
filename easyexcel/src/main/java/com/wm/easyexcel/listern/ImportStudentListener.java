package com.wm.easyexcel.listern;

import cn.hutool.core.collection.CollUtil;
import com.wm.easyexcel.entity.Student;
import com.wm.easyexcel.entity.Teacher;
import com.wm.easyexcel.service.IStudentService;
import com.wm.easyexcel.service.ITeacherService;
import com.wm.easyexcel.util.ReadExcelListener;
import org.springframework.stereotype.Component;

import javax.annotation.Resource;
import java.util.List;

/**
 * @ClassName: ImportStudentListener
 * @Description: 导入学生excel数据监听器
 * @Author: WM
 * @Date: 2022/12/22 10:44
 */
@Component
public class ImportStudentListener extends ReadExcelListener {

    @Resource
    private IStudentService studentService;
    @Resource
    private ITeacherService teacherService;

    /**
     * 实现数据保存逻辑
     *
     * @return
     */
    @Override
    public boolean saveData() {
        List dataList = getList();
        if (CollUtil.isEmpty(dataList)) {
            return false;
        }
        Object obj = dataList.get(0);
        if (obj instanceof Student) {
            return studentService.batchSaveStudent(dataList);
        } else if (obj instanceof Teacher) {
            return teacherService.batchSaveTeacher(dataList);
        }
        return true;
    }
}
