package com.wm.easyexcel.service;

import com.wm.easyexcel.entity.Student;

import java.util.List;
import java.util.Map;

/**
 * @ClassName: IStudentService
 * @Description: 学生相关接口
 * @Author: WM
 * @Date: 2023/1/9 16:52
 */
public interface IStudentService {

    /**
     * 查询学生集合信息
     *
     * @return
     */
    List<Student> listStudent(Long limit);

    /**
     * 批量保存学生数据
     *
     * @param students
     * @return
     */
    boolean batchSaveStudent(List<Student> students);

    /**
     * 批量保存学生数据-自行处理行列数据，不使用实体
     *
     * @param dataList
     * @param headMap
     * @return
     */
    boolean batchSaveStudent(List<Map<Integer, String>> dataList, Map<Integer, String> headMap);
}
