package com.easy_excel_demo;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.write.style.column.LongestMatchColumnWidthStyleStrategy;
import com.easy_excel_demo.demo.City;
import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;

import java.lang.reflect.Field;
import java.util.*;
import java.util.stream.Collectors;

@SpringBootTest
class EasyExcelDemoApplicationTests {

    @Test
    void normal() {
        List<City> data = City.demo();
        // 分组
        Map<String, List<City>> map = data.stream().collect(Collectors.groupingBy(City::getCityName));
        City last;
        for (String key : map.keySet()) {
            List<City> cur = map.get(key);
            last = new City(key, "小计");

            //  遍历添加
            cur.forEach(last::increase);

            cur.add(last);
        }

        for (List<City> cities : map.values())
            for (City cur : cities)
                System.out.println(cur);
    }

    @Test
    void dynamic() {
        String fileName = "E:\\动态表头.xlsx";

        EasyExcel.write(fileName, City.class)
                .sheet(0, "动态表头测试")
                .head(head())
                .registerWriteHandler(new LongestMatchColumnWidthStyleStrategy())
                .doWrite(Collections.emptyList());
    }

    public List<List<String>> head() {
        List<List<String>> result = new ArrayList<>();

        result.add(Arrays.asList("地市", "地市", "地市"));
        result.add(Arrays.asList("区县", "区县", "区县"));

        if (true) {
            result.add(Arrays.asList("故障工单", "应打卡", "应打卡"));
            result.add(Arrays.asList("故障工单", "已打卡", "质检合格"));
            result.add(Arrays.asList("故障工单", "已打卡", "质检不合格"));
            result.add(Arrays.asList("故障工单", "未打卡", "未打卡"));
        }

        if (false) {
            result.add(Arrays.asList("巡检工单", "应打卡", "应打卡"));
            result.add(Arrays.asList("巡检工单", "已打卡", "质检合格"));
            result.add(Arrays.asList("巡检工单", "已打卡", "质检不合格"));
            result.add(Arrays.asList("巡检工单", "未打卡", "未打卡"));
        }

        if (true) {
            result.add(Arrays.asList("发电工单", "应打卡", "应打卡"));
            result.add(Arrays.asList("发电工单", "已打卡", "质检合格"));
            result.add(Arrays.asList("发电工单", "已打卡", "质检不合格"));
            result.add(Arrays.asList("发电工单", "未打卡", "未打卡"));
        }

        if (false) {
            result.add(Arrays.asList("维修工单", "应打卡", "应打卡"));
            result.add(Arrays.asList("维修工单", "已打卡", "质检合格"));
            result.add(Arrays.asList("维修工单", "已打卡", "质检不合格"));
            result.add(Arrays.asList("维修工单", "未打卡", "未打卡"));
        }

        result.add(Arrays.asList("冗余", "冗余", "冗余"));

        return result;
    }

    @Test
    void test() {
        City city = new City("广州市", "白云区", null, null, null, 2L, 1L, 1L, 2L, 1L, 1L, 2L, 1L, 1L, 0L);

        try {
            Field field = city.getClass().getDeclaredField("faultSignAll");
            field.setAccessible(true);

            Object o = field.get(city);

            System.out.println(o);
        } catch (Exception e) {
            System.out.println(e);
        }
    }
}
