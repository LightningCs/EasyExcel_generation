package com.easy_excel_demo;

import com.easy_excel_demo.demo.City;
import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;

import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

@SpringBootTest
class EasyExcelDemoApplicationTests {

    @Test
    void contextLoads() {
        List<City> data = City.demo();
        // 分组
        Map<String, List<City>> map = data.stream().collect(Collectors.groupingBy(City::getCityName));
        City last;
        for (String key : map.keySet()) {
            List<City> cur = map.get(key);
            last = new City(key, "小计");

            //  遍历添加
            cur.forEach(last::add);

            cur.add(last);
        }

        for (List<City> cities : map.values())
            for (City cur : cities)
                System.out.println(cur);
    }

}
