package com.easy_excel_demo.style;

import com.alibaba.excel.write.metadata.style.WriteCellStyle;
import com.easy_excel_demo.strategy.ExcelStyleStrategy;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;

public class WriteStyle {

    public static ExcelStyleStrategy style() {
        // 创建头样式策略
        WriteCellStyle head = new WriteCellStyle();
        head.setVerticalAlignment(VerticalAlignment.CENTER); // 水平居中
        head.setHorizontalAlignment(HorizontalAlignment.CENTER); // 垂直居中
        // 背景颜色
        head.setFillForegroundColor(IndexedColors.WHITE.getIndex()); // 背景颜色

        WriteCellStyle content = new WriteCellStyle();
        // 边的样式
        content.setVerticalAlignment(VerticalAlignment.CENTER);
        content.setHorizontalAlignment(HorizontalAlignment.CENTER);
        content.setBorderBottom(BorderStyle.THIN); // 底部线条格式
        content.setBorderRight(BorderStyle.THIN); // 右边线条格式

        return new ExcelStyleStrategy(head, content);
    }
}
