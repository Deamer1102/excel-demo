package com.wm.easyexcel.service.impl;

import cn.hutool.core.lang.UUID;
import com.wm.easyexcel.entity.Teacher;
import com.wm.easyexcel.service.ITeacherService;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Service;

import java.util.ArrayList;
import java.util.List;

/**
 * @ClassName: StudentServiceImpl
 * @Description: 老师相关接口实现
 * @Author: WM
 * @Date: 2023/1/9 16:53
 */
@Slf4j
@Service
public class TeacherServiceImpl implements ITeacherService {

    @Override
    public List<Teacher> listTeacher(Long limit) {
        List<Teacher> teachers = new ArrayList<>();
        for (int i = 0; i < limit; i++) {
            Teacher teacher = new Teacher();
            teacher.setSex("女");
            teacher.setName("老师A" + i);
            teacher.setSno(UUID.fastUUID().toString());
            teachers.add(teacher);
        }
        return teachers;
    }

    @Override
    public boolean batchSaveTeacher(List<Teacher> teachers) {
        // 调用dao层接口保存数据...
        log.info("批量保存老师数据，共：{}条数据", teachers.size());
        return true;
    }

}
