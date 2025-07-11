package com.easy_excel_demo.controller;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.alibaba.excel.write.metadata.WriteTable;
import com.alibaba.excel.write.metadata.style.WriteCellStyle;
import com.easy_excel_demo.demo.City;
import com.easy_excel_demo.strategy.ExcelStyleStrategy;
import com.easy_excel_demo.writeHandle.ContentWriteHandler;
import com.easy_excel_demo.writeHandle.TitleRowWriteHandler;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

@RestController
@RequestMapping("/city")
@RequiredArgsConstructor
@Slf4j
public class CityController {

    @GetMapping
    public List<City> getCities() {

        // 读取
        EasyExcel.read();

        return null;
    }

    @GetMapping("/export")
    public void export() {
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

        Map<String, List<City>> map = City.demo().stream().collect(Collectors.groupingBy(City::getCityName));
        City last, end = new City("省公司", "合计");
        List<City> result = new ArrayList<>();
        // 加工添加小计
        for (String key : map.keySet()) {
            List<City> cur = map.get(key);
            // 清空数据
            last = new City(key, "小计");

            // 添加数据
            for (City item : map.get(key)) {
                last.add(item);
                end.add(item);
            }

            result.addAll(cur);
            result.add(last);
        }

        // 结尾
        result.add(end);

        // 保存到E盘
        String fileName ="E:\\双打卡报表.xlsx";
        String title = "“双打卡”预防监控(" + LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyy年MM月dd日HH时")) + ")";

        // 关闭资源
        try (ExcelWriter excelWriter = EasyExcel.write(fileName, City.class).build()) {
            WriteSheet sheet = EasyExcel.writerSheet("城市数据")
                    .needHead(Boolean.FALSE).build();

            WriteTable table0 = EasyExcel.writerTable(0)
                    .needHead(Boolean.TRUE)
                    .registerWriteHandler(new ExcelStyleStrategy(head, content))
                    .registerWriteHandler(new TitleRowWriteHandler(title))
                    .registerWriteHandler(new ContentWriteHandler())
                    .relativeHeadRowIndex(1)
                    .build();

            excelWriter.write(result, sheet, table0);
        }
    }
}
