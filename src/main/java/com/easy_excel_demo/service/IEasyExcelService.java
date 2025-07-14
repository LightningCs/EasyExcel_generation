package com.easy_excel_demo.service;

import java.util.List;
import java.util.function.Function;

public interface IEasyExcelService<T> {
    boolean export(Integer sheetNo, String sheetName, String fileName, String title, List<T> list, Function<List<T>, List<T>> f);
}
