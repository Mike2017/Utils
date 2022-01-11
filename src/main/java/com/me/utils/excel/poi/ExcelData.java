package com.me.utils.excel.poi;

import lombok.AllArgsConstructor;
import lombok.Data;

import java.util.List;

@Data
@AllArgsConstructor
public class ExcelData {

    /**
     * excel 名字
     */
    private String excelName;

    /**
     * 单个excel中的sheet数据
     */
    private List<SheetData> sheetDataList;
}
