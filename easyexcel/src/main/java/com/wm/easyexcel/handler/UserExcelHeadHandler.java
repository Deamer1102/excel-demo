package com.wm.easyexcel.handler;

import com.alibaba.excel.metadata.data.WriteCellData;
import com.alibaba.excel.write.handler.context.CellWriteHandlerContext;
import com.alibaba.excel.write.metadata.style.WriteCellStyle;
import com.alibaba.excel.write.metadata.style.WriteFont;
import com.alibaba.excel.write.style.HorizontalCellStyleStrategy;
import com.wm.easyexcel.util.ExcelHeadStyles;
import org.apache.poi.ss.usermodel.IndexedColors;

/**
 * @ClassName: UserExcelHeadHandler
 * @Description: 用户相关excel处理器
 * @Author: WM
 * @Date: 2023/1/9 18:24
 */
public class UserExcelHeadHandler extends HorizontalCellStyleStrategy {

    private ExcelHeadStyles excelHeadStyle;


    public UserExcelHeadHandler(ExcelHeadStyles excelHeadStyle) {
        this.excelHeadStyle = excelHeadStyle;
    }

    //设置头样式
    @Override
    protected void setHeadCellStyle(CellWriteHandlerContext context) {
        int columnNo = excelHeadStyle.getColumnIndex();
        WriteCellStyle headWriteCellStyle = new WriteCellStyle();
        headWriteCellStyle.setFillForegroundColor(IndexedColors.WHITE.getIndex());
        WriteFont headWriteFont = new WriteFont();
        headWriteFont.setFontHeightInPoints((short) 11);
        if (context.getColumnIndex() <= columnNo) {
            headWriteFont.setColor(excelHeadStyle.getFontColor());
        } else {
            headWriteFont.setColor(IndexedColors.BLACK.getIndex());
        }
        headWriteCellStyle.setWriteFont(headWriteFont);
        if (stopProcessing(context)) {
            return;
        }
        WriteCellData<?> cellData = context.getFirstCellData();
        WriteCellStyle.merge(headWriteCellStyle, cellData.getOrCreateStyle());
    }
}
