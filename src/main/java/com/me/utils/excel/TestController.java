package com.me.utils.excel;

import com.google.common.collect.Lists;
import com.me.utils.excel.ali.AliExcelData;
import com.me.utils.excel.ali.AliExcelUtils;
import com.me.utils.excel.ali.AliSheetData;
import com.me.utils.excel.ali.SheetTestModel;
import com.me.utils.excel.ali.ZipData;
import com.me.utils.excel.poi.ExcelData;
import com.me.utils.excel.poi.ExcelUtils;
import com.me.utils.excel.poi.SheetData;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.util.Date;
import java.util.List;

/**
 * 参考资料：
 * https://www.cnblogs.com/skyislimit/articles/10514719.html
 * https://blog.csdn.net/weixin_43525116/article/details/88063722
 * https://www.cnblogs.com/wushenghfut/p/11212106.html
 * https://blog.csdn.net/houxuehan/article/details/89189820
 * https://blog.csdn.net/xiaozhezhe0470/article/details/110877323
 */
@RestController
public class TestController {

//    public static void main(HttpServletRequest request, HttpServletResponse response) {
//        test1(request,response);
//        test2(request,response);
//
//    }

    @GetMapping("/exportByPoi")
    public void test1(HttpServletRequest request, HttpServletResponse response) {

        // 第一张表
        // 第一个sheet页面
        List<SheetData> sheetList1 = Lists.newArrayList();
        SheetData sheetData11 = new SheetData();
        sheetData11.setSheetName("test001");
        sheetData11.setTitles(Lists.newArrayList("姓名", "年龄", "地址"));
        List<List<Object>> dataList11 = Lists.newArrayList();
        List<Object> list11 = Lists.newArrayList("张三", "1", "北京");
        dataList11.add(list11);
        List<Object> list12 = Lists.newArrayList("李四", "2", "上海");
        dataList11.add(list12);
        sheetData11.setDataList(dataList11);
        sheetList1.add(sheetData11);
        // 第二个sheet页面
        SheetData sheetData22 = new SheetData();
        sheetData22.setSheetName("test002");
        sheetData22.setTitles(Lists.newArrayList("姓名", "班级", "总分"));
        List<List<Object>> dataList22 = Lists.newArrayList();
        List<Object> list21 = Lists.newArrayList("张三", "15", "500");
        dataList22.add(list21);
        List<Object> list22 = Lists.newArrayList("李四", "9", "600");
        dataList22.add(list22);
        sheetData22.setDataList(dataList22);
        sheetList1.add(sheetData22);

        // 第二张表
        // 第一个sheet页面
        List<SheetData> sheetList2 = Lists.newArrayList();
        SheetData sheetData33 = new SheetData();
        sheetData33.setSheetName("test003");
        sheetData33.setTitles(Lists.newArrayList("姓名", "年龄", "地址"));
        List<List<Object>> dataList33 = Lists.newArrayList();
        List<Object> list31 = Lists.newArrayList("张三2", "3", "武汉");
        dataList33.add(list31);
        List<Object> list32 = Lists.newArrayList("李四2", "4", "杭州");
        dataList33.add(list32);
        sheetData33.setDataList(dataList33);
        sheetList2.add(sheetData33);
        // 第二个sheet页面
        SheetData sheetData44 = new SheetData();
        sheetData44.setSheetName("test004");
        sheetData44.setTitles(Lists.newArrayList("姓名", "班级", "总分"));
        List<List<Object>> dataList44 = Lists.newArrayList();
        List<Object> list41 = Lists.newArrayList("张三2", "1", "100");
        dataList44.add(list41);
        List<Object> list42 = Lists.newArrayList("李四2", "2", "200");
        dataList44.add(list42);
        sheetData44.setDataList(dataList44);
        sheetList2.add(sheetData44);

        ExcelUtils.exportSingleExcel(request, response, new ExcelData("单表", sheetList1));

//        List<ExcelData> excelDataList = Lists.newArrayList( new ExcelData("表1",sheetList1),
//                new ExcelData("表2",sheetList2));
//        ExcelUtils.exportZipExcel(request, response, excelDataList);
    }

    @GetMapping("/exportByAli")
    public static void test2(HttpServletRequest request, HttpServletResponse response) {

        // 第一张表
        // 第一个sheet页面
        List<AliSheetData<SheetTestModel>> sheetList1 = Lists.newArrayList();
        AliSheetData<SheetTestModel> sheetData11 = new AliSheetData<>();
        sheetData11.setSheetName("test001");
        List<SheetTestModel> dataList11 = Lists.newArrayList();
        dataList11.add(new SheetTestModel("1",new Date(),1.0,"iii"));
        dataList11.add(new SheetTestModel("2",new Date(),2.0,"iii"));
        sheetData11.setDataList(dataList11);
        sheetList1.add(sheetData11);

        // 第二个sheet页面
        AliSheetData<SheetTestModel> sheetData22 = new AliSheetData<>();
        sheetData22.setSheetName("test002");
        List<SheetTestModel> dataList22 = Lists.newArrayList();
        dataList22.add(new SheetTestModel("3",new Date(),3.0,"iii"));
        dataList22.add(new SheetTestModel("4",new Date(),4.0,"iii"));
        sheetData22.setDataList(dataList22);
        sheetList1.add(sheetData22);

        // 第二张表
        // 第一个sheet页面
        List<AliSheetData<SheetTestModel>> sheetList2 = Lists.newArrayList();
        AliSheetData<SheetTestModel> sheetData33 = new AliSheetData<>();
        sheetData33.setSheetName("test003");
        List<SheetTestModel> dataList33 = Lists.newArrayList();
        dataList33.add(new SheetTestModel("5",new Date(),5.0,"iii"));
        dataList33.add(new SheetTestModel("6",new Date(),6.0,"iii"));
        sheetData33.setDataList(dataList33);
        sheetList2.add(sheetData33);
        // 第二个sheet页面
        AliSheetData<SheetTestModel> sheetData44 = new AliSheetData<>();
        sheetData44.setSheetName("test004");
        List<SheetTestModel> dataList44 = Lists.newArrayList();
        dataList44.add(new SheetTestModel("7",new Date(),7.0,"iii"));
        dataList44.add(new SheetTestModel("8",new Date(),8.0,"iii"));
        sheetData44.setDataList(dataList44);
        sheetList2.add(sheetData44);

        AliExcelData<SheetTestModel> excelData = new AliExcelData<>("表1", sheetList1,SheetTestModel.class);
//      AliExcelUtils.exportMultiSheet(response,excelData,AliSheetModel.class); // 泛型不能用在静态方法吗

        AliExcelData<SheetTestModel> excelData2 = new AliExcelData<>("表2", sheetList2,SheetTestModel.class);
        List<AliExcelData<SheetTestModel>> list = Lists.newArrayList();
        list.add(excelData);
        list.add(excelData2);
        AliExcelUtils.exportZipExcel(response, new ZipData<>("标签",list)); // 泛型不能用在静态方法吗
    }
}
