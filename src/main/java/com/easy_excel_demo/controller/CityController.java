package com.easy_excel_demo.controller;

import com.easy_excel_demo.demo.City;
import com.easy_excel_demo.service.IEasyExcelService;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;

@RestController
@RequestMapping("/city")
@RequiredArgsConstructor
@Slf4j
public class CityController {
    private final IEasyExcelService<City> easyExcelService;

    @GetMapping("/export")
    public void export() {
        boolean isSuccess = easyExcelService.export(0,
                "数据统计",
                "E:\\双打卡报表.xlsx",
                "“双打卡”预防监控(" + LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyy年MM月dd日HH时")),
                City.demo(),
                (list) -> {
                    List<City> result = new ArrayList<>();

                    City last = null, end = new City("省公司", "合计");
                    // 加工添加小计
                    for (City cur : list) {
                        if (last == null) {
                            last = new City(cur.getCityName(), "小计");
                        }

                        if (!last.getCityName().equals(cur.getCityName())) {
                            result.add(last);
                            last = null;
                        } else
                            last.increase(cur);

                        result.add(cur);
                        end.increase(cur);
                    }

                    result.add(last);
                    result.add(end);

                    return result;
                });

        if (isSuccess)
            log.info("导出成功!");
        else
            log.error("导出失败!");
    }
}
