package com.easy_excel_demo.controller;

import com.alibaba.excel.EasyExcel;
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
import java.util.*;
import java.util.stream.Collectors;

@RestController
@RequestMapping("/city")
@RequiredArgsConstructor
@Slf4j
public class CityController {

    @GetMapping("/export/check")
    public void export_check() {
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

//        Map<String, List<City>> map = City.demo().stream().collect(Collectors.groupingBy(City::getCityName));
        Map<String, List<City>> map = City.demo0().stream().collect(Collectors.groupingBy(City::getCityName));
//        Map<String, List<City>> map1 = City.demo1().stream().collect(Collectors.groupingBy(City::getCityName));
//        Map<String, List<City>> map2 = City.demo2().stream().collect(Collectors.groupingBy(City::getCityName));
//        Map<String, List<City>> map3 = City.demo3().stream().collect(Collectors.groupingBy(City::getCityName));

        List<City> result = map(map);

        // 保存到E盘
        String fileName ="E:\\双打卡报表1.xlsx";
        String title = "“双打卡”预防监控(" + LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyy年MM月dd日HH时")) + ")";

        // 关闭资源
        EasyExcel.write(fileName, City.class)
                .sheet(0, "城市数据")
                .head(head(result))
                .registerWriteHandler(new ExcelStyleStrategy(head, content))
                .registerWriteHandler(new TitleRowWriteHandler(title))
                .registerWriteHandler(new ContentWriteHandler())
                .relativeHeadRowIndex(1)
                .doWrite(data(result));
    }

    public List<List<String>> head(List<City> data) {
        List<List<String>> result = new ArrayList<>();
        City city = data.get(0);

        result.add(Arrays.asList("地市", "地市"));
        result.add(Arrays.asList("区县", "区县"));

        if (city.getFaultSignAll() != null) {
            result.add(Arrays.asList("故障工单", "应打卡"));
            result.add(Arrays.asList("故障工单", "已打卡"));
            result.add(Arrays.asList("故障工单", "未打卡"));
        }

        if (city.getInspectionSignAll() != null) {
            result.add(Arrays.asList("巡检工单", "应打卡"));
            result.add(Arrays.asList("巡检工单", "已打卡"));
            result.add(Arrays.asList("巡检工单", "未打卡"));
        }

        if (city.getGeneratorSignAll() != null) {
            result.add(Arrays.asList("发电工单", "应打卡"));
            result.add(Arrays.asList("发电工单", "已打卡"));
            result.add(Arrays.asList("发电工单", "未打卡"));
        }

        if (city.getMaintenanceSignAll() != null) {
            result.add(Arrays.asList("维修工单", "应打卡"));
            result.add(Arrays.asList("维修工单", "已打卡"));
            result.add(Arrays.asList("维修工单", "未打卡"));
        }

        result.add(Arrays.asList("冗余", "冗余"));

        return result;
    }

    public List<List<Object>> data(List<City> data) {
        List<List<Object>> result = new ArrayList<>();

        data.forEach((item) -> {
            List<Object> tmp = new ArrayList<>();

            tmp.add(item.getCityName());
            tmp.add(item.getCountyName());

            if (item.getFaultSignAll() != null) {
                tmp.add(item.getFaultSignAll());
                tmp.add(item.getFaultSignFinish());
                tmp.add(item.getFaultSignLack());
            }

            if (item.getInspectionSignAll() != null) {
                tmp.add(item.getInspectionSignAll());
                tmp.add(item.getInspectionSignFinish());
                tmp.add(item.getInspectionSignLack());
            }

            if (item.getGeneratorSignAll() != null) {
                tmp.add(item.getGeneratorSignAll());
                tmp.add(item.getGeneratorSignFinish());
                tmp.add(item.getGeneratorSignLack());
            }

            if (item.getMaintenanceSignAll() != null) {
                tmp.add(item.getMaintenanceSignAll());
                tmp.add(item.getMaintenanceSignFinish());
                tmp.add(item.getMaintenanceSignLack());
            }

            tmp.add(item.getOverflow());

            result.add(tmp);
        });

        return result;
    }

    public List<City> map(Map<String, List<City>> map) {
        List<City> result = new ArrayList<>();

        City last, end = new City("省公司", "合计");
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

        return result;
    }
}
