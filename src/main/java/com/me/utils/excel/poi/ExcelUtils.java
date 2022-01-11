package com.me.utils.excel.poi;

import com.me.utils.utils.DateUtils;
import com.me.utils.excel.ZipUtils;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.ByteArrayOutputStream;
import java.io.UnsupportedEncodingException;
import java.net.URLEncoder;
import java.util.List;
import java.util.Objects;
import java.util.zip.ZipOutputStream;

@Slf4j
public class ExcelUtils {

    private static final String DEFAULT_SHEET = "sheet1";

    /** 导出单excel，支持多sheet
     * @param request
     * @param response
     * @param excelData
     */
    public static void exportSingleExcel(HttpServletRequest request, HttpServletResponse response,
                                         ExcelData excelData) {
        SXSSFWorkbook workbook = null;
        try {
            String name = buildFileName(request, excelData.getExcelName() + ".xlsx");
            response.setCharacterEncoding("utf-8");
            response.addHeader("Content-Disposition", "attachment; filename=\"" + name + "\"");
            response.setContentType("application/octet-stream" + ";charset=" + "UTF-8");
            response.setHeader("Accept-Ranges", "bytes");
            workbook = buildWorkbook(excelData.getSheetDataList());
            workbook.write(response.getOutputStream());
        } catch (Exception e) {
            log.error("exportSingleExcel error", e);
        } finally {
            try {
                if (Objects.nonNull(workbook)) {
                    workbook.close();
                    workbook.dispose();
                }
            } catch (Exception e) {
                log.error("exportSingleExcel workbook close error", e);
            }
        }
    }

    private static String buildFileName(HttpServletRequest request, String name)
            throws UnsupportedEncodingException {
        String userAgent = request.getHeader("User-Agent");
        if (userAgent.contains("Firefox")) {
            return new String(name.getBytes("UTF-8"), "ISO8859-1");
        }
        return URLEncoder.encode(name, "UTF-8");
    }


    /**
     * 支持多excel，多sheet，打包zip包
     * @param request
     * @param response
     * @param excelDataList
     */
    public static void exportZipExcel(HttpServletRequest request, HttpServletResponse response,
                                      List<ExcelData> excelDataList) {
        ZipOutputStream zipOutputStream = null;
        try {
            response.setContentType("application/x-download");
            response.setCharacterEncoding("utf-8");
            String name = buildFileName(request, "标签" + DateUtils.getMomentTime() + ".zip");
            response.setHeader("Content-Disposition", "attachment;filename=\"" + name + "\"");
            zipOutputStream = new ZipOutputStream(response.getOutputStream());
            for (ExcelData excelData : excelDataList) {
                ByteArrayOutputStream boStream = new ByteArrayOutputStream();
                SXSSFWorkbook workbook = buildWorkbook(excelData.getSheetDataList());
                workbook.write(boStream);
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


    /**
     * @param sheetList
     * @return
     */
    private static SXSSFWorkbook buildWorkbook(List<SheetData> sheetList) {
        //wb对象
        SXSSFWorkbook wb = new SXSSFWorkbook();
        CellStyle titleStyle = initTitleStyle(wb);
        CellStyle dataStyle = initCellStyle(wb);

        for (int index = 0; index < sheetList.size(); index++) {
            SheetData sheetData = sheetList.get(index);
            //创建sheet对象
            Sheet sheet = wb.createSheet(buildSheetName(sheetData));
            //设置列默认的宽度
            sheet.setDefaultColumnWidth(15);
            //创建表头行
            initTitle(sheetData, titleStyle, sheet);
            // 设置单元格
            initCell(sheetData, dataStyle, sheet);
        }
        return wb;
    }

    private static String buildSheetName(SheetData sheetData) {
        String sheetName = sheetData.getSheetName();
        return Objects.isNull(sheetName) ? DEFAULT_SHEET : sheetName;
    }

    private static void initCell(SheetData sheetData, CellStyle dataStyle, Sheet sheet) {
        List<List<Object>> dataList = sheetData.getDataList();
        if (CollectionUtils.isEmpty(dataList)) {
            return;
        }

        for (int i = 0; i < dataList.size(); i++) {
            Row row = sheet.createRow(i + 1);
            List<Object> rowData = dataList.get(i);
            for (int j = 0; j < rowData.size(); j++) {
                Cell cell = row.createCell(j);
                cell.setCellStyle(dataStyle);
                Object obj = rowData.get(j);
                cell.setCellValue(Objects.isNull(obj) ? "-" : String.valueOf(obj));
            }
        }
    }

    private static void initTitle(SheetData sheetData, CellStyle titleStyle, Sheet sheet) {
        List<String> titles = sheetData.getTitles();
        if (CollectionUtils.isEmpty(titles)) {
            return;
        }
        Row rowHead = sheet.createRow(0);
        //设置表头行内容，可以在这里对表头设置一些样式，标红呀，加粗之类的
        for (int i = 0; i < titles.size(); i++) {
            Cell cellHead = rowHead.createCell(i);
            cellHead.setCellStyle(titleStyle);
            cellHead.setCellValue(titles.get(i));
        }
    }

    private static CellStyle initTitleStyle(Workbook wb) {
        CellStyle columnHeadStyle = initCellStyle(wb);
        Font columnHeadFont = wb.createFont();
        columnHeadFont.setFontHeightInPoints((short) 12);
        columnHeadFont.setBold(true);
        columnHeadStyle.setFont(columnHeadFont);
        columnHeadStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        columnHeadStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
        return columnHeadStyle;
    }

    private static CellStyle initCellStyle(Workbook wb) {
        CellStyle columnHeadStyle = wb.createCellStyle();
        DataFormat format = wb.createDataFormat();
        columnHeadStyle.setDataFormat(format.getFormat("@"));
        columnHeadStyle.setAlignment(HorizontalAlignment.CENTER);
        columnHeadStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        columnHeadStyle.setLocked(true);
        columnHeadStyle.setWrapText(true);
        columnHeadStyle.setLeftBorderColor(HSSFColor.HSSFColorPredefined.BLACK.getIndex());
        columnHeadStyle.setBorderLeft(BorderStyle.THIN);
        columnHeadStyle.setRightBorderColor(HSSFColor.HSSFColorPredefined.BLACK.getIndex());
        columnHeadStyle.setBorderRight(BorderStyle.THIN);
        columnHeadStyle.setBorderBottom(BorderStyle.THIN);
        columnHeadStyle.setBottomBorderColor(HSSFColor.HSSFColorPredefined.BLACK.getIndex());
        return columnHeadStyle;
    }
}
