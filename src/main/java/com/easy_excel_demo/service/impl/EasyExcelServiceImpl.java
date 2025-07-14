package com.easy_excel_demo.service.impl;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.write.style.column.LongestMatchColumnWidthStyleStrategy;
import com.easy_excel_demo.demo.City;
import com.easy_excel_demo.demo.ColumnConfig;
import com.easy_excel_demo.service.IEasyExcelService;
import com.easy_excel_demo.style.WriteStyle;
import com.easy_excel_demo.writeHandle.ContentWriteHandler;
import com.easy_excel_demo.writeHandle.TitleRowWriteHandler;
import lombok.RequiredArgsConstructor;
import org.springframework.stereotype.Service;

import java.lang.reflect.Field;
import java.util.*;
import java.util.function.Function;

@Service
@RequiredArgsConstructor
public class EasyExcelServiceImpl<T> implements IEasyExcelService<T> {
    private final ColumnConfig columnConfig;

    @Override
    public boolean export(Integer sheetNo, String sheetName, String fileName, String title, List<T> list, Function<List<T>, List<T>> f) {
        // 关闭资源
        try {
            EasyExcel.write(fileName, City.class)
                    .sheet(sheetNo, sheetName)
                    .excludeColumnFieldNames(excludeColumn(list.get(0), columnConfig.getColumns()))
                    .registerWriteHandler(WriteStyle.style())
                    .registerWriteHandler(new TitleRowWriteHandler(title, columnConfig))
                    .registerWriteHandler(new ContentWriteHandler())
                    .registerWriteHandler(new LongestMatchColumnWidthStyleStrategy())
                    .relativeHeadRowIndex(1)
                    .doWrite(f == null ? list : f.apply(list));
        } catch (Exception e) {
            e.printStackTrace();
            return false;
        }

        return true;
    }

    private Collection<String> excludeColumn(T data, List<String> fields) throws Exception {
        // 排除字段集合
        Set<String> excludeFields = new HashSet<>();

        // 遍历需要检查的字段集合
        for (String field : fields) {
            // 获取已声明字段-私有字段
            Field cur = data.getClass().getDeclaredField(field);

            // 暴力反射，解除私有限制
            cur.setAccessible(true);
            if (cur.get(data) == null) {
                excludeFields.add(field);
                columnConfig.setPace(columnConfig.getPace() + 1);
            }
        }

        return excludeFields;
    }
}
