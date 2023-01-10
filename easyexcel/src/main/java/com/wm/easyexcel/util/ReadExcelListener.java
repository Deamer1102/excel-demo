package com.wm.easyexcel.util;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.exception.ExcelDataConvertException;
import lombok.extern.slf4j.Slf4j;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * @ClassName: ReadExcelListener
 * @Description: 读取excel文件数据监听器
 * @Author: WM
 * @Date: 2022/12/21 17:00
 */
@Slf4j
public abstract class ReadExcelListener<T> extends AnalysisEventListener<T> {

    private static final int BATCH_COUNT = 10000;

    /**
     * 数据集
     */
    private final List<T> list = new ArrayList<>();

    /**
     * 表头数据
     */
    private Map<Integer, String> headMap = null;

    private Map<String, Object> params;

    // 获取数据集
    public List<T> getList() {
        return this.list;
    }

    // 获取表头数据
    public Map<Integer, String> getHeadMap() {
        return this.headMap;
    }

    /**
     * 每条数据都会进入
     *
     * @param object
     * @param analysisContext
     */
    @Override
    public void invoke(T object, AnalysisContext analysisContext) {
        this.list.add(object);
        // 达到BATCH_COUNT了，需要去存储一次数据库，防止数据几万条数据在内存，容易OOM
        if (list.size() >= BATCH_COUNT) {
            saveData();
            // 存储完成清理 list
            list.clear();
        }
    }

    /**
     * 数据解析完调用
     *
     * @param analysisContext
     */
    @Override
    public void doAfterAllAnalysed(AnalysisContext analysisContext) {
        // 这里也要保存数据，确保最后遗留的数据也存储到数据库
        saveData();
        log.info("所有数据解析完成！");
    }

    /**
     * 读取到的表头信息
     *
     * @param headMap
     * @param context
     */
    @Override
    public void invokeHeadMap(Map<Integer, String> headMap, AnalysisContext context) {
        this.headMap = headMap;
    }

    /**
     * 异常时调用
     *
     * @param exception：
     * @param context：
     * @throws Exception
     */
    @Override
    public void onException(Exception exception, AnalysisContext context) throws Exception {
        // 数据解析异常
        if (exception instanceof ExcelDataConvertException) {
            ExcelDataConvertException excelDataConvertException = (ExcelDataConvertException) exception;
            throw new RuntimeException("第" + excelDataConvertException.getRowIndex() + "行" + excelDataConvertException.getColumnIndex() + "列" + "数据解析异常");
        }
        // 其他异常...
    }

    /**
     * 数据存储到数据库
     */
    public abstract boolean saveData();
}
