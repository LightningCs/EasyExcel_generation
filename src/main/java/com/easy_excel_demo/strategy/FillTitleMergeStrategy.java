package com.easy_excel_demo.strategy;

import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.metadata.data.WriteCellData;
import com.alibaba.excel.write.handler.CellWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteTableHolder;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.List;

public class FillTitleMergeStrategy implements CellWriteHandler {

    @Override
    public void afterCellDispose(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, List<WriteCellData<?>> cellDataList, Cell cell, Head head, Integer relativeRowIndex, Boolean isHead) {
        // 当前行索引
        int rowIndex = cell.getRowIndex();
        // 当前列索引
        int columnIndex = cell.getColumnIndex();

        // 遍历单元格，判断单元格是否包含第三行子标题位置
        if (rowIndex == 2 && columnIndex > 1) {
            // 获取当前页
            Sheet sheet = writeSheetHolder.getSheet();

            List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();
            for (int i = 0; i < mergedRegions.size(); i++) {
                CellRangeAddress cellAddresses = mergedRegions.get(i);
                // 单元格是否已经合并
                if (cellAddresses.isInRange(rowIndex, columnIndex) && cellAddresses.isInRange(rowIndex, columnIndex + 1)) {
                    sheet.removeMergedRegion(i);
                    break;
                }
            }
        }
    }
}
