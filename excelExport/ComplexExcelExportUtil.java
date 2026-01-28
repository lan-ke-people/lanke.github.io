package com.qax.situation.asset.application.service.impl.excel.util;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.alibaba.excel.write.style.AbstractCellStyleStrategy;
import com.qax.needle.framework.boot.spring.MockMultipartFile;
import com.qax.situation.asset.application.dto.excel.export.UnitGroupDto;
import com.qax.situation.asset.application.service.impl.excel.builder.SystemExportDataBuilder;
import com.qax.situation.asset.application.service.impl.excel.strategy.ComplexHeaderMergeStrategy;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

/**
 * @author L-wangxinzhuo
 * @version 1.0
 * @description:
 * @date 2026/1/27 17:40
 */
@Slf4j
public class ComplexExcelExportUtil {

    /**
     * 导出Excel并直接返回MultipartFile
     * @param fileName 文件名（不含扩展名）
     * @param unitGroups 数据
     * @return MultipartFile
     */
    public static MultipartFile exportToMultipartFile(String fileName, List<UnitGroupDto> unitGroups) throws IOException {
        byte[] excelBytes = ComplexExcelExportUtil.exportComplexExcelToBytes(fileName, unitGroups);

        return new MockMultipartFile(
                "file", // form-data中的参数名
                fileName + ".xlsx", // 完整文件名
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", // MIME类型
                new ByteArrayInputStream(excelBytes)
        );
    }

    /**
     * 从临时文件创建MultipartFile
     */
    public static MultipartFile createMultipartFileFromTemp(String filePath) throws IOException {
        File file = new File(filePath);
        FileInputStream input = new FileInputStream(file);

        return new MockMultipartFile(
                "file",
                file.getName(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                input
        );
    }

    /**
     * 导出复杂结构的Excel到字节数组
     * @param fileName 文件名（不含扩展名）
     * @param unitGroups 数据
     * @return 字节数组
     */
    public static byte[] exportComplexExcelToBytes(String fileName, List<UnitGroupDto> unitGroups) throws IOException {
        log.info("=== 开始导出Excel ===");
        log.info("文件名：{}, 数据组数：{}", fileName, unitGroups == null ? 0 : unitGroups.size());

        if (unitGroups == null || unitGroups.isEmpty()) {
            log.error("unitGroups为空！");
            throw new IllegalArgumentException("导出数据不能为空");
        }

        try {
            // 创建ByteArrayOutputStream来接收Excel数据
            ByteArrayOutputStream outputStream = new ByteArrayOutputStream();

            // 使用outputStream创建ExcelWriter
            ExcelWriter excelWriter = EasyExcel.write(outputStream)
                    .registerWriteHandler(new ComplexHeaderMergeStrategy(unitGroups))
                    .registerWriteHandler(new CustomCellStyleHandler())
                    .build();

            // 构建数据
            List<List<Object>> testData = SystemExportDataBuilder.buildComplexData(unitGroups);

            WriteSheet testSheet = EasyExcel.writerSheet("重点保护对象清单").build();
            excelWriter.write(testData, testSheet);
            excelWriter.finish();

            // 通过outputStream获取字节数组，而不是通过excelWriter
            byte[] testBytes = outputStream.toByteArray();
            log.info("数据字节数组大小：{} 字节", testBytes.length);
            return testBytes;

        } catch (Exception e) {
            log.error("导出过程中出现异常：", e);
            throw new IOException("导出失败: " + e.getMessage(), e);
        }
    }

    /**
     * 创建简单的测试数据
     */
    private static List<List<Object>> createSimpleTestData() {
        List<List<Object>> data = new ArrayList<>();

        // 添加表头
        List<Object> header = new ArrayList<>();
        for (int i = 0; i < 35; i++) {
            header.add("列" + (i+1));
        }
        data.add(header);

        // 添加一行数据
        List<Object> row = new ArrayList<>();
        row.add("测试数据");
        for (int i = 1; i < 35; i++) {
            row.add("值" + i);
        }
        data.add(row);

        return data;
    }

    /**
     * 导出复杂结构的Excel到临时文件
     * @param fileName 文件名（不含扩展名）
     * @param unitGroups 数据
     * @return 临时文件路径
     */
    public static String exportComplexExcelToTempFile(String fileName, List<UnitGroupDto> unitGroups) throws IOException {
        byte[] bytes = exportComplexExcelToBytes(fileName, unitGroups);

        // 校验返回的数据是否为空
        if (bytes.length == 0) {
            throw new IOException("Excel导出失败，返回的字节数组为空");
        }

        // 创建临时文件
        String tempFileName = fileName + ".xlsx";
        String tempFilePath = System.getProperty("java.io.tmpdir") + File.separator + tempFileName;

        try (FileOutputStream fos = new FileOutputStream(tempFilePath)) {
            fos.write(bytes);
        }

        log.info("Excel已保存到临时文件：{}", tempFilePath);
        return tempFilePath;
    }

    /**
     * 自定义单元格样式处理器
     */
    public static class CustomCellStyleHandler extends AbstractCellStyleStrategy {

        @Override
        protected void setHeadCellStyle(Cell cell, Head head, Integer relativeRowIndex) {
            CellStyle cellStyle = cell.getSheet().getWorkbook().createCellStyle();

            // 设置字体
            Font font = cell.getSheet().getWorkbook().createFont();
            font.setBold(true);
            font.setFontHeightInPoints((short) 12);
            cellStyle.setFont(font);

            // 设置边框
            cellStyle.setBorderBottom(BorderStyle.THIN);
            cellStyle.setBorderLeft(BorderStyle.THIN);
            cellStyle.setBorderRight(BorderStyle.THIN);
            cellStyle.setBorderTop(BorderStyle.THIN);

            // 设置居中对齐
            cellStyle.setAlignment(HorizontalAlignment.CENTER);
            cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);

            // 设置背景色（表头行）
            if (relativeRowIndex <= 2) {
                cellStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
                cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            }

            cell.setCellStyle(cellStyle);
        }

        @Override
        protected void setContentCellStyle(Cell cell, Head head, Integer relativeRowIndex) {
            CellStyle cellStyle = cell.getSheet().getWorkbook().createCellStyle();

            // 设置边框
            cellStyle.setBorderBottom(BorderStyle.THIN);
            cellStyle.setBorderLeft(BorderStyle.THIN);
            cellStyle.setBorderRight(BorderStyle.THIN);
            cellStyle.setBorderTop(BorderStyle.THIN);

            // 设置居中对齐
            cellStyle.setAlignment(HorizontalAlignment.CENTER);
            cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);

            // 设置自动换行
            cellStyle.setWrapText(true);

            cell.setCellStyle(cellStyle);
        }
    }
}
