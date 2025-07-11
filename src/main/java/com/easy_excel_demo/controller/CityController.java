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
        List<City> list = City.init();

        // 创建样式策略
        WriteCellStyle head = new WriteCellStyle();
        head.setVerticalAlignment(VerticalAlignment.CENTER);
        head.setHorizontalAlignment(HorizontalAlignment.CENTER);
        // 背景颜色
        head.setFillForegroundColor(IndexedColors.WHITE.getIndex());

        WriteCellStyle content = new WriteCellStyle();
        // 边的样式
        content.setVerticalAlignment(VerticalAlignment.CENTER);
        content.setHorizontalAlignment(HorizontalAlignment.CENTER);
        content.setBorderBottom(BorderStyle.THIN);
        content.setBorderRight(BorderStyle.THIN);

        List<City> data = processed(City.init());

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

            excelWriter.write(data, sheet, table0);
        }
    }

    /**
     * 添加 小计和总计
     * @param data
     * @return data
     */
    public List<City> processed(List<City> data) {
        City temp = new City("小计"), all = new City("合计");
        all.setCityName("省公司");
        ArrayList<City> newData = new ArrayList<>(data);
        int len = data.size(), j = 0;

        // j 偏移量
        for (int i = 0; i < len; i++) {
            City city = data.get(i);

            // 当前模板城市名不为空且前后不一致
            if (temp.getCityName() != null && city.getCityName() != temp.getCityName()) {
                newData.add(i + j, temp);
                temp = new City("小计");
                j++;
                continue;
            }

            temp.setCityName(city.getCityName());
            temp.setFaultSignAll(temp.getFaultSignAll() + city.getFaultSignAll());
            temp.setFaultSignFinish(temp.getFaultSignFinish() + city.getFaultSignFinish());
            temp.setFaultSignLack(temp.getFaultSignLack() + city.getFaultSignLack());
            temp.setInspectionSignAll(temp.getInspectionSignAll() + city.getInspectionSignAll());
            temp.setInspectionSignFinish(temp.getInspectionSignFinish() + city.getInspectionSignFinish());
            temp.setInspectionSignLack(temp.getInspectionSignLack() + city.getInspectionSignLack());
            temp.setGeneratorSignAll(temp.getGeneratorSignAll() + city.getGeneratorSignAll());
            temp.setGeneratorSignFinish(temp.getGeneratorSignFinish() + city.getGeneratorSignFinish());
            temp.setGeneratorSignLack(temp.getGeneratorSignLack() + city.getGeneratorSignLack());
            temp.setMaintenanceSignAll(temp.getMaintenanceSignAll() + city.getMaintenanceSignAll());
            temp.setMaintenanceSignFinish(temp.getMaintenanceSignFinish() + city.getMaintenanceSignFinish());
            temp.setMaintenanceSignLack(temp.getMaintenanceSignLack() + city.getMaintenanceSignLack());
            temp.setOverflow(temp.getOverflow() + city.getOverflow());

            all.setFaultSignAll(all.getFaultSignAll() + city.getFaultSignAll());
            all.setFaultSignFinish(all.getFaultSignFinish() + city.getFaultSignFinish());
            all.setFaultSignLack(all.getFaultSignLack() + city.getFaultSignLack());
            all.setInspectionSignAll(all.getInspectionSignAll() + city.getInspectionSignAll());
            all.setInspectionSignFinish(all.getInspectionSignFinish() + city.getInspectionSignFinish());
            all.setInspectionSignLack(all.getInspectionSignLack() + city.getInspectionSignLack());
            all.setGeneratorSignAll(all.getGeneratorSignAll() + city.getGeneratorSignAll());
            all.setGeneratorSignFinish(all.getGeneratorSignFinish() + city.getGeneratorSignFinish());
            all.setGeneratorSignLack(all.getGeneratorSignLack() + city.getGeneratorSignLack());
            all.setMaintenanceSignAll(all.getMaintenanceSignAll() + city.getMaintenanceSignAll());
            all.setMaintenanceSignFinish(all.getMaintenanceSignFinish() + city.getMaintenanceSignFinish());
            all.setMaintenanceSignLack(all.getMaintenanceSignLack() + city.getMaintenanceSignLack());
            all.setOverflow(all.getOverflow() + city.getOverflow());
        }

        newData.add(len + j++, temp);
        newData.add(len + j, all);

        return newData;
    }
}
