package com.wm.easyexcel.service.impl;

import cn.hutool.core.lang.UUID;
import com.wm.easyexcel.entity.Student;
import com.wm.easyexcel.service.IStudentService;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Service;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * @ClassName: StudentServiceImpl
 * @Description: 学生相关接口实现
 * @Author: WM
 * @Date: 2023/1/9 16:53
 */
@Slf4j
@Service
public class StudentServiceImpl implements IStudentService {

    @Override
    public List<Student> listStudent(Long limit) {
        List<Student> students = new ArrayList<>();
        for (int i = 0; i < limit; i++) {
            Student student = new Student();
            student.setSex("男");
            student.setNum(i);
            student.setName("小A" + i);
            student.setSNo(UUID.fastUUID().toString());
            students.add(student);
        }
        return students;
    }

    @Override
    public boolean batchSaveStudent(List<Student> students) {
        // 调用dao层接口保存数据...
        log.info("批量保存学生数据，共：{}条数据", students.size());
        return true;
    }

    @Override
    public boolean batchSaveStudent(List<Map<Integer, String>> dataList, Map<Integer, String> headMap) {
        // 这里可以针对读取到的表头数据和对应的行数据，这种应用场景适用于导入的excel中数据需要我们自己解析做其他业务处理，或者表头列是动态的，我们
        // 没有办法使用一个固定的实体做映射
        log.info("表头列数据：{}", headMap.values());
        log.info("批量保存学生数据，共：{}条数据", dataList.size());
        return true;
    }
}
