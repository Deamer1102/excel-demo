package com.wm.easyexcel.service;

import com.wm.easyexcel.entity.Teacher;

import java.util.List;
import java.util.Map;

/**
 * @ClassName: ITeacherService
 * @Description: 老师相关接口
 * @Author: WM
 * @Date: 2023/1/9 16:52
 */
public interface ITeacherService {

    /**
     * 查询老师集合信息
     *
     * @return
     */
    List<Teacher> listTeacher(Long limit);

    /**
     * 批量保存老师数据
     *
     * @param Teachers
     * @return
     */
    boolean batchSaveTeacher(List<Teacher> Teachers);

}
