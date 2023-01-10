package com.wm.easyexcel.util;

import cn.hutool.core.collection.CollUtil;
import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.EasyExcelFactory;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.read.builder.ExcelReaderBuilder;
import com.alibaba.excel.util.StringUtils;
import com.alibaba.excel.write.builder.ExcelWriterSheetBuilder;
import com.alibaba.excel.write.handler.WriteHandler;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.alibaba.excel.write.metadata.fill.FillConfig;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;

/**
 * @ClassName: EasyExcelUtil
 * @Description: excel导入导出工具类
 * @Author: WM
 * @Date: 2023/1/9 15:36
 */
public class EasyExcelUtil {

    /**
     * 导入-读取一页sheet
     *
     * @param inputStream 文件流
     * @param clazz       数据对象
     * @param sheetName   要读取的sheet (不传，默认读取第一个sheet)
     * @throws Exception
     */
    public static <T> List<T> importExcel(InputStream inputStream, Class<T> clazz, String sheetName, ReadExcelListener readExcelListener) throws Exception {
        ExcelReaderBuilder builder = EasyExcelFactory.read(inputStream, clazz, readExcelListener);
        if (StringUtils.isBlank(sheetName)) {
            builder.sheet().doRead();
        } else {
            builder.sheet(sheetName).doRead();
        }
        return readExcelListener.getList();
    }

    /**
     * 导入-读取多个sheet
     *
     * @param inputStream  文件流
     * @param sheetNum     需要读取的sheet个数(默认0开始，如果传入2，则读取0、1)
     * @param sheetObjList 每个sheet里面需要封装的对象(如果index为2，则需要传入对应的2个对象)
     * @param <T>
     * @return
     */
    public static <T> List<List<T>> importExcels(InputStream inputStream, int sheetNum, List<T> sheetObjList, ReadExcelListener<T> readExcelListener) throws Exception {
        List<List<T>> resultList = new LinkedList<>();
        for (int index = 0; index < sheetNum; index++) {
            EasyExcelFactory.read(inputStream, sheetObjList.get(index).getClass(), readExcelListener).sheet(index).doRead();
            List<T> list = readExcelListener.getList();
            resultList.add(list);
        }
        return resultList;
    }

    /**
     * 导出excel表格;只导出模板
     *
     * @param response
     * @param headList     表头列表
     * @param dataList     数据列表
     * @param fileName     文件名称
     * @param sheetName    sheet名称
     * @param writeHandler 处理器
     * @throws Exception
     */
    public static <T> void exportExcel(HttpServletResponse response, List<List<String>> headList, List<List<T>> dataList, String fileName, String sheetName, WriteHandler writeHandler) throws Exception {
        response.setContentType("application/vnd.ms-excel");
        response.setCharacterEncoding(StandardCharsets.UTF_8.name());
        fileName = URLEncoder.encode(fileName, StandardCharsets.UTF_8.name());
        response.setHeader("Content-disposition", "attachment;filename=" + fileName + FileTypeEnum.TEMPLATE_SUFFIX.getDesc());
        ExcelWriter excelWriter = EasyExcel.write(response.getOutputStream()).build();
        // 这里注意如果同一个sheet只要创建一次
        ExcelWriterSheetBuilder sheetBuilder = EasyExcel.writerSheet(sheetName).head(headList);
        if (null != writeHandler) {
            sheetBuilder.registerWriteHandler(writeHandler);
        }
        WriteSheet writeSheet = sheetBuilder.build();
        excelWriter.write(dataList, writeSheet);
        excelWriter.finish();
    }

    /**
     * 导出excel表格-列不确定情况;只导出模板
     *
     * @param response
     * @param headList      表头列表
     * @param dataList      数据列表
     * @param fileName      文件名称
     * @param sheetName     sheet名称
     * @param writeHandlers 处理器集合
     * @throws Exception
     */
    public static <T> void exportExcel(HttpServletResponse response, List<List<String>> headList, List<List<T>> dataList, String fileName, String sheetName, List<WriteHandler> writeHandlers) throws Exception {
        response.setContentType("application/vnd.ms-excel");
        response.setCharacterEncoding(StandardCharsets.UTF_8.name());
        fileName = URLEncoder.encode(fileName, StandardCharsets.UTF_8.name());
        response.setHeader("Content-disposition", "attachment;filename=" + fileName + FileTypeEnum.TEMPLATE_SUFFIX.getDesc());
        ExcelWriter excelWriter = EasyExcel.write(response.getOutputStream()).build();
        // 这里注意如果同一个sheet只要创建一次
        ExcelWriterSheetBuilder sheetBuilder = EasyExcel.writerSheet(sheetName).head(headList);
        if (CollUtil.isNotEmpty(writeHandlers)) {
            for (WriteHandler writeHandler : writeHandlers) {
                sheetBuilder.registerWriteHandler(writeHandler);
            }
        }
        WriteSheet writeSheet = sheetBuilder.build();
        excelWriter.write(dataList, writeSheet);
        excelWriter.finish();
    }

    /**
     * 导出excel表格
     *
     * @param response
     * @param dataList  数据列表
     * @param clazz     数据对象
     * @param fileName  文件名称
     * @param sheetName sheet名称
     * @throws Exception
     */
    public static <T> void exportExcel(HttpServletResponse response, List<T> dataList, Class<T> clazz, String fileName, String sheetName) throws Exception {
        response.setContentType("application/vnd.ms-excel");
        response.setCharacterEncoding(StandardCharsets.UTF_8.name());
        fileName = URLEncoder.encode(fileName, StandardCharsets.UTF_8.name());
        response.setHeader("Content-disposition", "attachment;filename=" + fileName + FileTypeEnum.TEMPLATE_SUFFIX.getDesc());
        EasyExcelFactory.write(response.getOutputStream(), clazz).sheet(sheetName).doWrite(dataList);
    }

    /**
     * 导出excel表格-支持行单元格合并
     *
     * @param response
     * @param dataList     数据列表
     * @param clazz        数据对象
     * @param fileName     文件名称
     * @param sheetName    sheet名称
     * @param writeHandler 处理器
     * @throws Exception
     */
    public static <T> void exportWriteHandlerExcel(HttpServletResponse response, List<T> dataList, Class<T> clazz, String fileName, String sheetName, WriteHandler writeHandler) throws Exception {
        response.setContentType("application/vnd.ms-excel");
        response.setCharacterEncoding(StandardCharsets.UTF_8.name());
        fileName = URLEncoder.encode(fileName, StandardCharsets.UTF_8.name());
        response.setHeader("Content-disposition", "attachment;filename=" + fileName + FileTypeEnum.TEMPLATE_SUFFIX.getDesc());
        ExcelWriter excelWriter = EasyExcel.write(response.getOutputStream(), clazz).build();
        // 这里注意如果同一个sheet只要创建一次
        WriteSheet writeSheet = EasyExcel.writerSheet(sheetName)
                .registerWriteHandler(writeHandler).build();
        excelWriter.write(dataList, writeSheet);
        excelWriter.finish();
    }


    /**
     * 导出多个sheet
     *
     * @param response
     * @param dataList 多个sheet数据列表
     * @param clazzMap 对应每个sheet列表里面的数据对应的sheet名称
     * @param fileName 文件名
     * @param <T>
     * @throws Exception
     */
    public static <T> void exportExcels(HttpServletResponse response, List<List<?>> dataList, Map<Integer, String> clazzMap, String fileName) throws Exception {
        response.setContentType("application/vnd.ms-excel");
        response.setCharacterEncoding(StandardCharsets.UTF_8.name());
        fileName = URLEncoder.encode(fileName, StandardCharsets.UTF_8.name());
        response.setHeader("Content-disposition", "attachment;filename=" + fileName + FileTypeEnum.TEMPLATE_SUFFIX.getDesc());
        ExcelWriter excelWriter = EasyExcel.write(response.getOutputStream()).build();
        int sheetNum = dataList.size();
        for (int i = 0; i < sheetNum; i++) {
            List<?> objects = dataList.get(i);
            Class<?> aClass = objects.get(0).getClass();
            WriteSheet writeSheet0 = EasyExcel.writerSheet(i, clazzMap.get(i)).head(aClass).build();
            excelWriter.write(objects, writeSheet0);
        }
        excelWriter.finish();
    }

    /**
     * 根据模板将集合对象填充表格-单个sheet
     *
     * @param inputStream 模板文件输入流
     * @param list        填充对象集合
     * @param object      填充对象
     * @param fileName    文件名称
     * @throws Exception
     */
    public static <T> void exportSheetTemplateExcel(InputStream inputStream, HttpServletResponse response, List<T> list, Object object, String fileName) throws Exception {
        FillConfig fillConfig = FillConfig.builder().forceNewRow(Boolean.TRUE).build();
        ExcelWriter excelWriter = EasyExcelFactory.write(getOutputStream(fileName, response)).withTemplate(inputStream).build();
        WriteSheet writeSheet0 = EasyExcelFactory.writerSheet(0).build();
        excelWriter.fill(object, fillConfig, writeSheet0);
        excelWriter.fill(list, fillConfig, writeSheet0);
        excelWriter.finish();
    }

    /**
     * 根据模板将集合对象填充表格-多个sheet
     *
     * @param inputStream 模板文件输入
     * @param response    模板文件输出流
     * @param list1       填充对象集合
     * @param list2       填充对象集合
     * @param object1     填充对象
     * @param object2     填充对象
     * @param fileName    文件名称
     * @throws Exception
     */
    public static <T> void exportSheetsTemplateExcel(InputStream inputStream, HttpServletResponse response, List<T> list1, List<T> list2, Object object1, Object object2, String fileName) throws Exception {
        FillConfig fillConfig = FillConfig.builder().forceNewRow(Boolean.TRUE).build();
        ExcelWriter excelWriter = EasyExcelFactory.write(getOutputStream(fileName, response)).withTemplate(inputStream).build();
        WriteSheet writeSheet0 = EasyExcelFactory.writerSheet(0).build();
        WriteSheet writeSheet1 = EasyExcelFactory.writerSheet(1).build();
        excelWriter.fill(object1, fillConfig, writeSheet0);
        excelWriter.fill(list1, fillConfig, writeSheet0);
        excelWriter.fill(object2, fillConfig, writeSheet1);
        excelWriter.fill(list2, fillConfig, writeSheet1);
        excelWriter.finish();
    }

    /**
     * 根据模板将单个对象填充表格
     *
     * @param inputStream 模板文件输入流
     * @param response    模板文件输出流
     * @param object      填充对象
     * @param fileName    文件名称
     * @param sheetName   需要写入的sheet名称(不传：填充到第一个sheet)
     * @throws Exception
     */
    public static void exportTemplateExcel(InputStream inputStream, HttpServletResponse response, Object object, String fileName, String sheetName) throws Exception {
        if (StringUtils.isBlank(sheetName)) {
            EasyExcelFactory.write(getOutputStream(fileName, response)).withTemplate(inputStream).sheet().doFill(object);
        } else {
            EasyExcelFactory.write(getOutputStream(fileName, response)).withTemplate(inputStream).sheet(sheetName).doFill(object);
        }
    }

    /**
     * 根据模板将集合对象填充表格
     *
     * @param inputStream 模板文件输入流
     * @param response    模板文件输出流
     * @param list        填充对象集合
     * @param fileName    文件名称
     * @param sheetName   需要写入的sheet(不传：填充到第一个sheet)
     * @throws Exception
     */
    public static <T> void exportTemplateExcelList(InputStream inputStream, HttpServletResponse response, List<T> list, String fileName, String sheetName) throws Exception {
        // 全部填充：全部加载到内存中一次填充
        if (StringUtils.isBlank(sheetName)) {
            EasyExcelFactory.write(getOutputStream(fileName, response)).withTemplate(inputStream).sheet().doFill(list);
        } else {
            EasyExcelFactory.write(getOutputStream(fileName, response)).withTemplate(inputStream).sheet(sheetName).doFill(list);
        }
    }

    /**
     * 构建输出流
     *
     * @param fileName 文件名称
     * @param response 模板文件输出流
     * @param response
     * @return
     * @throws Exception
     */
    private static OutputStream getOutputStream(String fileName, HttpServletResponse response) throws Exception {
        fileName = URLEncoder.encode(fileName, StandardCharsets.UTF_8.name());
        response.setContentType("application/vnd.ms-excel");
        response.setCharacterEncoding(StandardCharsets.UTF_8.name());
        response.setHeader("Content-Disposition", "attachment;filename=" + fileName + FileTypeEnum.TEMPLATE_SUFFIX.getDesc());
        return response.getOutputStream();
    }

    /**
     * 文件格式校验
     *
     * @param file
     */
    public static void checkFile(MultipartFile file) {
        if (file == null) {
            throw new RuntimeException("文件不能为空！");
        }
        String fileName = file.getOriginalFilename();
        if (StringUtils.isEmpty(fileName)) {
            throw new RuntimeException("文件名称不能为空！");
        }
        if (!fileName.endsWith(FileTypeEnum.TEMPLATE_SUFFIX.getDesc())
                && !fileName.endsWith(FileTypeEnum.TEMPLATE_SUFFIX_XLS.getDesc())) {
            throw new RuntimeException("请上传.xlsx或.xls文件！");
        }
    }
}
