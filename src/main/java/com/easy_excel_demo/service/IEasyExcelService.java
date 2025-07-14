package com.easy_excel_demo.service;

import java.util.List;

public interface IEasyExcelService<T> {
    boolean export(Integer sheetNo, String sheetName, String fileName, String title, List<T> list);
}
