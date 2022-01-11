package com.me.utils.excel.ali;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.me.utils.excel.ZipUtils;
import lombok.extern.slf4j.Slf4j;

import javax.servlet.http.HttpServletResponse;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.io.UnsupportedEncodingException;
import java.net.URLEncoder;
import java.util.List;
import java.util.Objects;
import java.util.zip.ZipOutputStream;

/**
 * 参考：https://github.com/alibaba/easyexcel
 * https://www.yuque.com/easyexcel/doc/write
 */
@Slf4j
public class AliExcelUtils {

    /**
     *  导出单excel，支持多sheet
     * @param response
     * @param excelData
     * @param <T>
     * @throws IOException
     */
    public static <T> void exportMultiSheet(HttpServletResponse response, AliExcelData<T> excelData) throws IOException {
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        response.setCharacterEncoding("utf-8");
        String fileName = getFileName(excelData.getExcelName());
        response.setHeader("Content-disposition", "attachment;filename*=utf-8''" + fileName + ".xlsx");
        multiSheet(response.getOutputStream(), excelData);
    }

    /**
     *  导出多excel的zip包
     *  每个excel支持多sheet
     * @param response
     * @param zipData
     * @param <T>
     */
    public static <T> void exportZipExcel(HttpServletResponse response, ZipData<T> zipData) {
        ZipOutputStream zipOutputStream = null;
        try {
            response.setContentType("application/x-download");
            response.setCharacterEncoding("utf-8");
            String fileName = getFileName(zipData.getZipName());
            response.setHeader("Content-Disposition", "attachment;filename*=utf-8''" + fileName + ".zip");
            zipOutputStream = new ZipOutputStream(response.getOutputStream());
            for (AliExcelData excelData : zipData.getExcelDataList()) {
                ByteArrayOutputStream boStream = new ByteArrayOutputStream();
                multiSheet(boStream, excelData);
                ZipUtils.compressFileToZipStream(zipOutputStream, boStream, excelData.getExcelName() + ".xlsx");
                boStream.close();
            }
        } catch (Exception e) {
            log.error("exportZipExcel error", e);
        } finally {
            try {
                if (Objects.nonNull(zipOutputStream)) {
                    zipOutputStream.flush();
                    zipOutputStream.close();
                }
            } catch (Exception e) {
                log.error("exportZipExcel zipOutputStream close error", e);
            }
        }
    }

    private static String getFileName(String name) throws UnsupportedEncodingException {
        return URLEncoder.encode(name, "UTF-8").replaceAll("\\+", "%20");
    }

    private static <T> void multiSheet(OutputStream outputStream, AliExcelData<T> excelData) {
        ExcelWriter excelWriter = null;
        try {
            // 指定输出到outputStream、定义数据模型
            excelWriter = EasyExcel.write(outputStream, excelData.getSheetDataClass()).build();
            List<AliSheetData<T>> sheetDataList = excelData.getSheetDataList();
            for (int i = 0; i < sheetDataList.size(); i++) {
                AliSheetData<T> sheetData = sheetDataList.get(i);
                WriteSheet writeSheet = EasyExcel.writerSheet(i, sheetData.getSheetName()).build();
                excelWriter.write(sheetData.getDataList(), writeSheet);
            }
        } finally {
            if (excelWriter != null) {
                excelWriter.finish();
            }
        }
    }
}
