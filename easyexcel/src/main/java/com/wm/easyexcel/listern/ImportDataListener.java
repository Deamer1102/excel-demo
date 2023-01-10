package com.wm.easyexcel.listern;

import cn.hutool.core.collection.CollUtil;
import com.wm.easyexcel.service.IStudentService;
import com.wm.easyexcel.util.ReadExcelListener;
import org.springframework.stereotype.Component;

import javax.annotation.Resource;
import java.util.List;
import java.util.Map;

/**
 * @ClassName: ImportDataListener
 * @Description: 导入数据监听器
 * @Author: WM
 * @Date: 2023/1/9 19:27
 */
@Component
public class ImportDataListener extends ReadExcelListener {

    @Resource
    private IStudentService studentService;

    /**
     * 实现数据保存逻辑-直接操作表头和行数据
     *
     * @return
     */
    @Override
    public boolean saveData() {
        Map<Integer, String> headMap = getHeadMap();
        List dataList = getList();
        if (CollUtil.isEmpty(headMap) || CollUtil.isEmpty(dataList)) {
            return false;
        }
        return studentService.batchSaveStudent(dataList, headMap);
    }
}
