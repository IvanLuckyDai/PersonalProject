package com.dai.Service;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.List;

public interface ExcelService {
    void initExcel();

    void getAllHead();

    void writeAllHeadToNewExcel(XSSFWorkbook new_workbook, List<String> heads);
}
