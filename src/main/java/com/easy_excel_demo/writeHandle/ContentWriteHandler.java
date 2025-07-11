package com.easy_excel_demo.writeHandle;

import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.metadata.data.WriteCellData;
import com.alibaba.excel.write.handler.CellWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteTableHolder;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.List;

@Slf4j
public class ContentWriteHandler implements CellWriteHandler {
    @Override
    public void afterCellDispose(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, List<WriteCellData<?>> cellDataList, Cell cell, Head head, Integer relativeRowIndex, Boolean isHead) {
        if (isHead) {
            log.info("\r\n当前行为表头，无需操作");
            return;
        }

        // 判断当前 Cell 所在列索引不是0不执行操作
        if (cell.getColumnIndex() != 0) {
            log.info("\r\n当前不是0列，无需操作");
            return;
        }

        // 当前sheet
        Sheet sheet = cell.getSheet();

        // 当前cell所在行索引
        int rowIndexCur = cell.getRowIndex();

        // 当前cell所在行的上一行索引
        int rowIndexPrev = rowIndexCur - 1;

        // 当前cell的row对象
        Row rowCur = cell.getRow();

        // 上一行的row对象
        Row rowPrev = sheet.getRow(rowIndexPrev);

        // 上一行cell
        Cell cellPrev = rowPrev.getCell(cell.getColumnIndex());

        // 获取上一行单元格和当前单元格的值
        Object cur = cell.getCellType() == CellType.STRING ? cell.getStringCellValue() : cell.getNumericCellValue();
        Object pre = cellPrev.getCellType() == CellType.STRING ? cellPrev.getStringCellValue() : cellPrev.getNumericCellValue();

        log.info("当前单元格的值是{}，上一行单元格的值是{}", cur, pre);

        // 判断是否一致
        if (!cur.equals(pre)) {
            // 不一致不进行操作
            log.info("\r\n两行单元格值不一致，无需操作");
            return;
        }

        log.info("\r\n两行单元格值一致，进行合并");

        // 获取所有合并区域
        List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();
        // 是否合并过
        boolean isMerge = false;

        for (int i = 0; i < mergedRegions.size(); i++) {
            CellRangeAddress cellAddresses = mergedRegions.get(i);

            if (cellAddresses.isInRange(rowIndexPrev, cell.getColumnIndex())) {
                sheet.removeMergedRegion(i);

                cellAddresses.setLastRow(rowIndexCur);
                sheet.addMergedRegion(cellAddresses);

                isMerge = true;
                break;
            }
        }

        if (!isMerge) {
            CellRangeAddress cellAddresses = new CellRangeAddress(rowIndexPrev, rowIndexCur, cell.getColumnIndex(), cell.getColumnIndex());
            sheet.addMergedRegion(cellAddresses);
        }

        log.info("\r\n合并单元格完成");
    }
}
