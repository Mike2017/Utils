package com.me.utils.excel.ali;

import lombok.Data;

import java.util.List;

@Data
public class AliSheetData<T> {

    /**
     * sheet名字
     */
    private String sheetName;

    /**
     * 单个sheet表对应的数据
     */
    private List<T> dataList;

}
