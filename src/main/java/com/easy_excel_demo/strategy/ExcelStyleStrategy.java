package com.easy_excel_demo.strategy;

import com.alibaba.excel.write.metadata.style.WriteCellStyle;
import com.alibaba.excel.write.style.HorizontalCellStyleStrategy;
import lombok.Data;

public class ExcelStyleStrategy extends HorizontalCellStyleStrategy {

    public ExcelStyleStrategy(WriteCellStyle head, WriteCellStyle content) {
        super(head, content);
    }
}
