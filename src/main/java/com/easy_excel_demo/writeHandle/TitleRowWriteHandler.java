package com.easy_excel_demo.writeHandle;

import com.alibaba.excel.write.handler.SheetWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteWorkbookHolder;
import com.easy_excel_demo.demo.City;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

@Data
@AllArgsConstructor
@NoArgsConstructor
public class TitleRowWriteHandler implements SheetWriteHandler {
    private String title;

    @Override
    public void afterSheetCreate(WriteWorkbookHolder writeWorkbookHolder, WriteSheetHolder writeSheetHolder) {
        Sheet sheet = writeSheetHolder.getSheet();
        Workbook workbook = writeSheetHolder.getParentWriteWorkbookHolder().getWorkbook();

        // 创建首行（索引0）作为标题行
        Row titleRow = sheet.createRow(0);
        Cell titleCell = titleRow.createCell(0);
        titleCell.setCellValue(title);

        // 设置标题样式（自定义）
        CellStyle titleStyle = workbook.createCellStyle();
        titleStyle.setAlignment(HorizontalAlignment.CENTER);
        titleStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        // 设置标题字体
        Font titleFont = workbook.createFont();
        titleFont.setBold(true);
        titleFont.setFontHeightInPoints((short) 15);
        titleFont.setFontName("宋体");
        titleStyle.setFont(titleFont);
        titleCell.setCellStyle(titleStyle);

        // 合并单元格（跨整行）
        int lastCol = City.class.getDeclaredFields().length - 1; // 根据实际列数调整
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, lastCol));
    }
}
