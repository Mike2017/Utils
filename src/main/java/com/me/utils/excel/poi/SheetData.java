package com.me.utils.excel.poi;

import lombok.Data;

import java.util.List;

@Data
public class SheetData {

    /**
     * sheet名字
     */
    private String sheetName;

    /**
     * 标题
     */
    private List<String> titles;

    /**
     * 单个sheet表对应的数据
     */
    private List<List<Object>> dataList;

}
