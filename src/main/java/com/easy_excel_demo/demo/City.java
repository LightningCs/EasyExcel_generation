package com.easy_excel_demo.demo;

import com.alibaba.excel.annotation.ExcelIgnore;
import com.alibaba.excel.annotation.ExcelIgnoreUnannotated;
import com.alibaba.excel.annotation.ExcelProperty;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.RequiredArgsConstructor;

import java.util.Arrays;
import java.util.List;

@Data
@AllArgsConstructor
@RequiredArgsConstructor
@ExcelIgnoreUnannotated
public class City {
    @ExcelProperty("地市")
//    @ExcelIgnore
//    @ColumnWidth(10)
    private String cityName;
    @ExcelProperty("区县")
//    @ColumnWidth(10)
    private String countyName;
    @ExcelProperty({"故障工单", "应打卡"})
//    @ColumnWidth(10)
    private Long faultSignAll;
    @ExcelProperty({"故障工单", "已打卡"})
//    @ColumnWidth(10)
    private Long faultSignFinish;
    @ExcelProperty({"故障工单", "未打卡"})
//    @ColumnWidth(10)
    private Long faultSignLack;
    @ExcelProperty({"巡检工单", "应打卡"})
//    @ColumnWidth(10)
    private Long inspectionSignAll;
    @ExcelProperty({"巡检工单", "已打卡"})
//    @ColumnWidth(10)
    private Long inspectionSignFinish;
    @ExcelProperty({"巡检工单", "未打卡"})
//    @ColumnWidth(10)
    private Long inspectionSignLack;
    @ExcelProperty({"发电工单", "应打卡"})
//    @ColumnWidth(10)
    private Long generatorSignAll;
    @ExcelProperty({"发电工单", "已打卡"})
//    @ColumnWidth(10)
    private Long generatorSignFinish;
    @ExcelProperty({"发电工单", "未打卡"})
//    @ColumnWidth(10)
    private Long generatorSignLack;
    @ExcelProperty({"维护工单", "应打卡"})
//    @ColumnWidth(10)
    private Long maintenanceSignAll;
    @ExcelProperty({"维护工单", "已打卡"})
//    @ColumnWidth(10)
    private Long maintenanceSignFinish;
    @ExcelProperty({"维护工单", "未打卡"})
//    @ColumnWidth(10)
    private Long maintenanceSignLack;
    @ExcelProperty("冗余")
//    @ColumnWidth(10)
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
        this.faultSignAll = null;
        this.faultSignFinish = null;
        this.faultSignLack = null;
        this.inspectionSignAll = null;
        this.inspectionSignFinish = null;
        this.inspectionSignLack = null;
        this.generatorSignAll = null;
        this.generatorSignFinish = null;
        this.generatorSignLack = null;
        this.maintenanceSignAll = null;
        this.maintenanceSignFinish = null;
        this.maintenanceSignLack = null;
        this.overflow = null;
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

    public static List<City> demo0() {
        List<City> list = Arrays.asList(
                new City("广州市", "白云区", null, null, null, 2L, 1L, 1L, 2L, 1L, 1L, 2L, 1L, 1L, 0L),
                new City("广州市", "从化区", null, null, null, 2L, 1L, 1L, 2L, 1L, 1L, 2L, 1L, 1L, 0L),
                new City("广州市", "花都区", null, null, null, 2L, 1L, 1L, 2L, 1L, 1L, 2L, 1L, 1L, 0L),
                new City("广州市", "越秀区", null, null, null, 2L, 1L, 1L, 2L, 1L, 1L, 2L, 1L, 1L, 0L),
                new City("梅州市", "丰顺区", null, null, null, 2L, 1L, 1L, 2L, 1L, 1L, 2L, 1L, 1L, 0L),
                new City("梅州市", "蕉岭区", null, null, null, 2L, 1L, 1L, 2L, 1L, 1L, 2L, 1L, 1L, 0L)
        );

        return list;
    }

    public static List<City> demo1() {
        List<City> list = Arrays.asList(
                new City("广州市", "白云区", 2L, 1L, 1L, null, null, null, 2L, 1L, 1L, 2L, 1L, 1L, 0L),
                new City("广州市", "从化区", 2L, 1L, 1L, null, null, null, 2L, 1L, 1L, 2L, 1L, 1L, 0L),
                new City("广州市", "花都区", 2L, 1L, 1L, null, null, null, 2L, 1L, 1L, 2L, 1L, 1L, 0L),
                new City("广州市", "越秀区", 2L, 1L, 1L, null, null, null, 2L, 1L, 1L, 2L, 1L, 1L, 0L),
                new City("梅州市", "丰顺区", 2L, 1L, 1L, null, null, null, 2L, 1L, 1L, 2L, 1L, 1L, 0L),
                new City("梅州市", "蕉岭区", 2L, 1L, 1L, null, null, null, 2L, 1L, 1L, 2L, 1L, 1L, 0L)
        );

        return list;
    }

    public static List<City> demo2() {
        List<City> list = Arrays.asList(
                new City("广州市", "白云区", 2L, 1L, 1L, 2L, 1L, 1L, null, null, null, 2L, 1L, 1L, 0L),
                new City("广州市", "从化区", 2L, 1L, 1L, 2L, 1L, 1L, null, null, null, 2L, 1L, 1L, 0L),
                new City("广州市", "花都区", 2L, 1L, 1L, 2L, 1L, 1L, null, null, null, 2L, 1L, 1L, 0L),
                new City("广州市", "越秀区", 2L, 1L, 1L, 2L, 1L, 1L, null, null, null, 2L, 1L, 1L, 0L),
                new City("梅州市", "丰顺区", 2L, 1L, 1L, 2L, 1L, 1L, null, null, null, 2L, 1L, 1L, 0L),
                new City("梅州市", "蕉岭区", 2L, 1L, 1L, 2L, 1L, 1L, null, null, null, 2L, 1L, 1L, 0L)
        );

        return list;
    }

    public static List<City> demo3() {
        List<City> list = Arrays.asList(
                new City("广州市", "白云区", 2L, 1L, 1L, 2L, 1L, 1L, 2L, 1L, 1L, null, null, null, 0L),
                new City("广州市", "从化区", 2L, 1L, 1L, 2L, 1L, 1L, 2L, 1L, 1L, null, null, null, 0L),
                new City("广州市", "花都区", 2L, 1L, 1L, 2L, 1L, 1L, 2L, 1L, 1L, null, null, null, 0L),
                new City("广州市", "越秀区", 2L, 1L, 1L, 2L, 1L, 1L, 2L, 1L, 1L, null, null, null, 0L),
                new City("梅州市", "丰顺区", 2L, 1L, 1L, 2L, 1L, 1L, 2L, 1L, 1L, null, null, null, 0L),
                new City("梅州市", "蕉岭区", 2L, 1L, 1L, 2L, 1L, 1L, 2L, 1L, 1L, null, null, null, 0L)
        );

        return list;
    }

    /**
     * 累加数据
     * @param item
     */
    public void add(City item) {
        if (item.getFaultSignAll() != null)
            this.faultSignAll = (this.faultSignAll == null ? 0 : this.faultSignAll) + item.getFaultSignAll();
        if (item.getFaultSignFinish() != null)
            this.faultSignFinish = (this.faultSignFinish == null ? 0 : this.faultSignFinish) + item.getFaultSignFinish();
        if (item.getFaultSignLack() != null)
            this.faultSignLack = (this.faultSignLack == null ? 0 : this.faultSignLack) + item.getFaultSignLack();
        if (item.getInspectionSignAll() != null)
            this.inspectionSignAll = (this.inspectionSignAll == null ? 0 : this.inspectionSignAll) + item.getInspectionSignAll();
        if (item.getInspectionSignFinish() != null)
            this.inspectionSignFinish = (this.inspectionSignFinish == null ? 0 : this.inspectionSignFinish) + item.getInspectionSignFinish();
        if (item.getInspectionSignLack() != null)
            this.inspectionSignLack = (this.inspectionSignLack == null ? 0 : this.inspectionSignLack) + item.getInspectionSignLack();
        if (item.getGeneratorSignAll() != null)
            this.generatorSignAll = (this.generatorSignAll == null ? 0 : this.generatorSignAll) + item.getGeneratorSignAll();
        if (item.getGeneratorSignFinish() != null)
            this.generatorSignFinish = (this.generatorSignFinish == null ? 0 : this.generatorSignFinish) + item.getGeneratorSignFinish();
        if (item.getGeneratorSignLack() != null)
            this.generatorSignLack = (this.generatorSignLack == null ? 0 : this.generatorSignLack) + item.getGeneratorSignLack();
        if (item.getMaintenanceSignAll() != null)
            this.maintenanceSignAll = (this.maintenanceSignAll == null ? 0 : this.maintenanceSignAll) + item.getMaintenanceSignAll();
        if (item.getMaintenanceSignFinish() != null)
            this.maintenanceSignFinish = (this.maintenanceSignFinish == null ? 0 : this.maintenanceSignFinish) + item.getMaintenanceSignFinish();
        if (item.getMaintenanceSignLack() != null)
            this.maintenanceSignLack = (this.maintenanceSignLack == null ? 0 : this.maintenanceSignLack) + item.getMaintenanceSignLack();
        this.overflow = (this.overflow == null ? 0 : this.overflow) + item.getOverflow();
    }
}
