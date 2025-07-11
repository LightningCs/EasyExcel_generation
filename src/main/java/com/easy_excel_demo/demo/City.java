package com.easy_excel_demo.demo;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.write.style.ColumnWidth;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.RequiredArgsConstructor;

import java.util.Arrays;
import java.util.List;

@Data
@AllArgsConstructor
@RequiredArgsConstructor
public class City {
    @ExcelProperty("地市")
    @ColumnWidth(10)
    private String cityName;
    @ExcelProperty("区县")
    @ColumnWidth(10)
    private String countyName;
    @ExcelProperty({"故障工单", "应打卡"})
    @ColumnWidth(10)
    private Long faultSignAll;
    @ExcelProperty({"故障工单", "已打卡"})
    @ColumnWidth(10)
    private Long faultSignFinish;
    @ExcelProperty({"故障工单", "未打卡"})
    @ColumnWidth(10)
    private Long faultSignLack;
    @ExcelProperty({"巡检工单", "应打卡"})
    @ColumnWidth(10)
    private Long inspectionSignAll;
    @ExcelProperty({"巡检工单", "已打卡"})
    @ColumnWidth(10)
    private Long inspectionSignFinish;
    @ExcelProperty({"巡检工单", "未打卡"})
    @ColumnWidth(10)
    private Long inspectionSignLack;
    @ExcelProperty({"发电工单", "应打卡"})
    @ColumnWidth(10)
    private Long generatorSignAll;
    @ExcelProperty({"发电工单", "已打卡"})
    @ColumnWidth(10)
    private Long generatorSignFinish;
    @ExcelProperty({"发电工单", "未打卡"})
    @ColumnWidth(10)
    private Long generatorSignLack;
    @ExcelProperty({"维护工单", "应打卡"})
    @ColumnWidth(10)
    private Long maintenanceSignAll;
    @ExcelProperty({"维护工单", "已打卡"})
    @ColumnWidth(10)
    private Long maintenanceSignFinish;
    @ExcelProperty({"维护工单", "未打卡"})
    @ColumnWidth(10)
    private Long maintenanceSignLack;
    @ExcelProperty("冗余")
    @ColumnWidth(10)
    private Long overflow;

    public City(String countyName) {
        this.countyName = countyName;
        init();
    }

    public City(String cityName, String countyName) {
        this.cityName = cityName;
        this.countyName = countyName;
        init();
    }

    /**
     * 初始化数据
     */
    public void init() {
        this.faultSignAll = 0L;
        this.faultSignFinish = 0L;
        this.faultSignLack = 0L;
        this.inspectionSignAll = 0L;
        this.inspectionSignFinish = 0L;
        this.inspectionSignLack = 0L;
        this.generatorSignAll = 0L;
        this.generatorSignFinish = 0L;
        this.generatorSignLack = 0L;
        this.maintenanceSignAll = 0L;
        this.maintenanceSignFinish = 0L;
        this.maintenanceSignLack = 0L;
        this.overflow = 0L;
    }

    public static List<City> demo() {
        List<City> list = Arrays.asList(
                new City("广州市", "白云区", 2L, 1L, 1L, 2L, 1L, 1L, 2L, 1L, 1L, 2L, 1L, 1L, 0L),
                new City("广州市", "从化区", 2L, 1L, 1L, 2L, 1L, 1L, 2L, 1L, 1L, 2L, 1L, 1L, 0L),
                new City("广州市", "花都区", 2L, 1L, 1L, 2L, 1L, 1L, 2L, 1L, 1L, 2L, 1L, 1L, 0L),
                new City("广州市", "越秀区", 2L, 1L, 1L, 2L, 1L, 1L, 2L, 1L, 1L, 2L, 1L, 1L, 0L),
                new City("梅州市", "丰顺区", 2L, 1L, 1L, 2L, 1L, 1L, 2L, 1L, 1L, 2L, 1L, 1L, 0L),
                new City("梅州市", "蕉岭区", 2L, 1L, 1L, 2L, 1L, 1L, 2L, 1L, 1L, 2L, 1L, 1L, 0L)
        );

        return list;
    }

    /**
     * 累加数据
     * @param item
     */
    public void add(City item) {
        this.faultSignAll += item.getFaultSignAll();
        this.faultSignFinish += item.getFaultSignFinish();
        this.faultSignLack += item.getFaultSignLack();
        this.inspectionSignAll += item.getInspectionSignAll();
        this.inspectionSignFinish += item.getInspectionSignFinish();
        this.inspectionSignLack += item.getInspectionSignLack();
        this.generatorSignAll += item.getGeneratorSignAll();
        this.generatorSignFinish += item.getGeneratorSignFinish();
        this.generatorSignLack += item.getGeneratorSignLack();
        this.maintenanceSignAll += item.getMaintenanceSignAll();
        this.maintenanceSignFinish += item.getMaintenanceSignFinish();
        this.maintenanceSignLack += item.getMaintenanceSignLack();
        this.overflow += item.getOverflow();
    }
}
