package com.me.utils.excel.ali;

import lombok.AllArgsConstructor;
import lombok.Data;

import java.util.List;

@Data
@AllArgsConstructor
public class ZipData<T> {

    /**
     * 压缩包名字
     */
    private String zipName;

    /**
     * 多个excel数据
     */
    private List<AliExcelData<T>> excelDataList;
}
