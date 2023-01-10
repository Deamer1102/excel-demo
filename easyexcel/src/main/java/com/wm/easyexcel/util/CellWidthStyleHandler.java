package com.wm.easyexcel.util;

import cn.hutool.core.collection.CollUtil;
import com.alibaba.excel.enums.CellDataTypeEnum;
import com.alibaba.excel.metadata.data.CellData;
import com.alibaba.excel.metadata.data.WriteCellData;
import com.alibaba.excel.write.handler.context.CellWriteHandlerContext;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.style.column.AbstractColumnWidthStyleStrategy;
import org.apache.poi.ss.usermodel.Cell;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @ClassName: CellWidthStyleHandler
 * @Description: 设置表头的调整宽策略
 * @Author: WM
 * @Date: 2022/12/23 14:37
 */
public class CellWidthStyleHandler extends AbstractColumnWidthStyleStrategy {

    // 可以根据这里的最大宽度，按自己需要进行调整,搭配单元格样式实现类中的，自动换行，效果更好
    private static final int MAX_COLUMN_WIDTH = 50;
    private Map<Integer, Map<Integer, Integer>> CACHE = new HashMap(8);

    /**
     * 是否设置固定宽度
     */
    private boolean fixed;

    /**
     * 固定宽度
     */
    private int fixedWidth;

    public CellWidthStyleHandler() {
    }

    public CellWidthStyleHandler(boolean fixed, int fixedWidth) {
        this.fixed = fixed;
        this.fixedWidth = (fixedWidth == 0 ? 15 : fixedWidth);
    }

    @Override
    protected void setColumnWidth(CellWriteHandlerContext context) {
        boolean needSetWidth = context.getHead() || !CollUtil.isEmpty(context.getCellDataList());
        if (!needSetWidth) {
            return;
        }
        WriteSheetHolder writeSheetHolder = context.getWriteSheetHolder();
        // 设置固定宽度
        if (fixed) {
            Cell cell = context.getCell();
            writeSheetHolder.getSheet().setColumnWidth(cell.getColumnIndex(), fixedWidth * 256);
            return;
        }
        // 设置自动调整宽度
        Map<Integer, Integer> maxColumnWidthMap = CACHE.get(writeSheetHolder.getSheetNo());
        if (maxColumnWidthMap == null) {
            maxColumnWidthMap = new HashMap(16);
            CACHE.put(writeSheetHolder.getSheetNo(), maxColumnWidthMap);
        }
        Cell cell = context.getCell();
        int columnWidth = this.dataLength(context.getCellDataList(), cell, context.getHead());
        if (columnWidth >= 0) {
            if (columnWidth > MAX_COLUMN_WIDTH) {
                columnWidth = MAX_COLUMN_WIDTH;
            }
            Integer maxColumnWidth = maxColumnWidthMap.get(cell.getColumnIndex());
            if (maxColumnWidth == null || columnWidth > maxColumnWidth) {
                maxColumnWidthMap.put(cell.getColumnIndex(), columnWidth);
                writeSheetHolder.getSheet().setColumnWidth(cell.getColumnIndex(), columnWidth * 256);
            }
        }
    }

    private int dataLength(List<WriteCellData<?>> cellDataList, Cell cell, Boolean isHead) {
        if (isHead) {
            return cell.getStringCellValue().getBytes().length;
        } else {
            CellData cellData = cellDataList.get(0);
            CellDataTypeEnum type = cellData.getType();
            if (type == null) {
                return -1;
            } else {
                switch (type) {
                    case STRING:
                        return cellData.getStringValue().getBytes().length;
                    case BOOLEAN:
                        return cellData.getBooleanValue().toString().getBytes().length;
                    case NUMBER:
                        return cellData.getNumberValue().toString().getBytes().length;
                    default:
                        return -1;
                }
            }
        }
    }
}
