package com.me.utils.excel.ali;

import lombok.AllArgsConstructor;
import lombok.Data;

import java.util.List;

@Data
@AllArgsConstructor
public class AliExcelData<T> {

    /**
     * excel 名字
     */
    private String excelName;

    /**
     * 单个excel中的sheet数据
     */
    private List<AliSheetData<T>> sheetDataList;

    /**
     * 模型class类
     */
    private Class<T> sheetDataClass;
}
