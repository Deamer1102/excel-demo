package com.wm.easyexcel.util;

import com.alibaba.excel.metadata.data.WriteCellData;
import com.alibaba.excel.write.handler.context.CellWriteHandlerContext;
import com.alibaba.excel.write.metadata.style.WriteCellStyle;
import com.alibaba.excel.write.metadata.style.WriteFont;
import com.alibaba.excel.write.style.HorizontalCellStyleStrategy;
import org.apache.poi.ss.usermodel.IndexedColors;

/**
 * @ClassName: DefaultExcelHeadHandler
 * @Description: 默认自定义表头样式拦截器
 * @Author: WM
 * @Date: 2022/12/22 17:35
 */
public class DefaultExcelHeadHandler extends HorizontalCellStyleStrategy {

    //设置头样式
    @Override
    protected void setHeadCellStyle(CellWriteHandlerContext context) {
        WriteCellStyle headWriteCellStyle = new WriteCellStyle();
        headWriteCellStyle.setFillForegroundColor(IndexedColors.WHITE.getIndex());
        WriteFont headWriteFont = new WriteFont();
        headWriteFont.setFontHeightInPoints((short) 11);
        headWriteFont.setColor(IndexedColors.BLACK.getIndex());
        headWriteCellStyle.setWriteFont(headWriteFont);
        if (stopProcessing(context)) {
            return;
        }
        WriteCellData<?> cellData = context.getFirstCellData();
        WriteCellStyle.merge(headWriteCellStyle, cellData.getOrCreateStyle());
    }
}
