package cn.itcast.common.excel.utils;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;

/**
 * Created by liwang on 2017/9/8.
 */
public class ExcelStyle {
    // === Excel文件的头部标题样式
    public static CellStyle getHeaderStyle(Workbook workbook){
        CellStyle headerStyle = workbook.createCellStyle();
        // === 设置边框
        headerStyle.setBorderBottom(BorderStyle.THIN);
        headerStyle.setBorderLeft(BorderStyle.THIN);
        headerStyle.setBorderRight(BorderStyle.THIN);
        headerStyle.setBorderTop(BorderStyle.THIN);

        // === 设置背景色
        headerStyle.setFillForegroundColor(HSSFColor.HSSFColorPredefined.LIGHT_GREEN.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        // === 设置居中
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        // === 设置字体
        Font font = workbook.createFont();
        font.setFontName("粗体");

        // === 设置字体大小
        font.setFontHeightInPoints((short) 16);

        // === 设置粗体显示
        font.setBold(true);

        // === 选择需要用到的字体格式
        headerStyle.setFont(font);
        return headerStyle;
    };
    /*
      * 设置Excel文件的第二列的注意事项提示信息的样式
      */
    private static CellStyle setWarnerCellStyles(Workbook workbook) {
        CellStyle warnerStyle = workbook.createCellStyle();

        // === 设置边框
        warnerStyle.setBorderBottom(BorderStyle.THIN);
        warnerStyle.setBorderLeft(BorderStyle.THIN);
        warnerStyle.setBorderRight(BorderStyle.THIN);
        warnerStyle.setBorderTop(BorderStyle.THIN);

        // === 设置背景色
        warnerStyle.setFillForegroundColor(HSSFColor.HSSFColorPredefined.LIGHT_GREEN.getIndex());
        warnerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        // === 设置左对齐
        warnerStyle.setAlignment(HorizontalAlignment.LEFT);

        // === 设置字体
        Font font = workbook.createFont();
        font.setFontName("宋体");

        // === 设置字体大小
        font.setFontHeightInPoints((short) 10);

        // === 设置粗体显示
        font.setBold(true);

        // === 设置字体颜色
        font.setColor(HSSFColor.HSSFColorPredefined.RED.getIndex());

        // === 选择需要用到的字体格式
        warnerStyle.setFont(font);

        // === 设置自动换行
        warnerStyle.setWrapText(true);
        return warnerStyle;
    }

    /*
     * 设置Excel文件的列头样式
     */
    public static CellStyle setTitleCellStyles(Workbook workbook) {
        CellStyle titleStyle = workbook.createCellStyle();
        // === 设置边框
        titleStyle.setBorderBottom(BorderStyle.THIN);
        titleStyle.setBorderLeft(BorderStyle.THIN);
        titleStyle.setBorderRight(BorderStyle.THIN);
        titleStyle.setBorderTop(BorderStyle.THIN);

        // === 设置背景色
        titleStyle.setFillForegroundColor(HSSFColor.HSSFColorPredefined.LIGHT_GREEN.getIndex());
        titleStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        // === 设置居中
        titleStyle.setAlignment(HorizontalAlignment.CENTER);
        titleStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        // === 设置字体
        Font font = workbook.createFont();
        font.setFontName("粗体");

        // === 设置字体大小
        font.setFontHeightInPoints((short) 12);

        // === 设置粗体显示
        font.setBold(true);

        // === 选择需要用到的字体格式
        titleStyle.setFont(font);

        // === 设置自动换行
        // titleStyle.setWrapText(true) ;
        return titleStyle;
    }

    /*
     * 设置Excel文件的数据样式
     */
    public static CellStyle setDataCellStyles(Workbook workbook) {
        CellStyle dataStyle = workbook.createCellStyle();
        // === 设置单元格格式为文本格式
        DataFormat dataFormat = workbook.createDataFormat();
        dataStyle.setDataFormat(dataFormat.getFormat("@"));
        dataStyle.setBorderBottom(BorderStyle.THIN);
        dataStyle.setBorderLeft(BorderStyle.THIN);
        dataStyle.setBorderRight(BorderStyle.THIN);
        dataStyle.setBorderTop(BorderStyle.THIN);

        // === 设置背景色
        dataStyle.setFillForegroundColor(HSSFColor.HSSFColorPredefined.LIGHT_GREEN.getIndex());
        dataStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        // === 设置居中
        dataStyle.setAlignment(HorizontalAlignment.CENTER);
        dataStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        // === 设置字体
        Font font = workbook.createFont();
        font.setFontName("宋体");

        // === 设置字体大小
        font.setFontHeightInPoints((short) 11);

        // === 选择需要用到的字体格式
        dataStyle.setFont(font);

        // === 设置自动换行
        // dataStyle.setWrapText(true) ;
        return dataStyle;
    }


    public static CellStyle setNumberDataCellStyles(Workbook workbook) {
        CellStyle dataStyle = workbook.createCellStyle();
        // === 设置单元格格式为文本格式
        DataFormat dataFormat = workbook.createDataFormat();
        dataStyle.setDataFormat(dataFormat.getFormat("#,#0.00"));
        dataStyle.setBorderBottom(BorderStyle.THIN);
        dataStyle.setBorderLeft(BorderStyle.THIN);
        dataStyle.setBorderRight(BorderStyle.THIN);
        dataStyle.setBorderTop(BorderStyle.THIN);

        // === 设置背景色
        dataStyle.setFillForegroundColor(HSSFColor.HSSFColorPredefined.LIGHT_GREEN.getIndex());
        dataStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        // === 设置居中
        dataStyle.setAlignment(HorizontalAlignment.CENTER);
        dataStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        // === 设置字体
        Font font = workbook.createFont();
        font.setFontName("宋体");

        // === 设置字体大小
        font.setFontHeightInPoints((short) 11);

        // === 选择需要用到的字体格式
        dataStyle.setFont(font);

        // === 设置自动换行
        // dataStyle.setWrapText(true) ;
        return dataStyle;
    }

    /*
     * 错误数据重新导入Excel中的样式
     */
    private static CellStyle setErrorDataStyle(Workbook workbook) {
        CellStyle errorDataStyle = workbook.createCellStyle();
        // === 设置边框颜色
        errorDataStyle.setBottomBorderColor(HSSFColor.HSSFColorPredefined.RED.getIndex());
        errorDataStyle.setLeftBorderColor(HSSFColor.HSSFColorPredefined.RED.getIndex());
        errorDataStyle.setRightBorderColor(HSSFColor.HSSFColorPredefined.RED.getIndex());
        errorDataStyle.setTopBorderColor(HSSFColor.HSSFColorPredefined.RED.getIndex());

        // === 设置边框
        errorDataStyle.setBorderBottom(BorderStyle.THIN);
        errorDataStyle.setBorderLeft(BorderStyle.THIN);
        errorDataStyle.setBorderRight(BorderStyle.THIN);
        errorDataStyle.setBorderTop(BorderStyle.THIN);

        // === 设置背景色
        errorDataStyle.setFillForegroundColor(HSSFColor.HSSFColorPredefined.ROSE.getIndex());
        errorDataStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        // === 设置居中
        errorDataStyle.setAlignment(HorizontalAlignment.LEFT);
        errorDataStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        // === 设置字体
        Font font = workbook.createFont();
        font.setFontName("宋体");

        // === 设置字体大小
        font.setFontHeightInPoints((short) 11);

        // === 选择需要用到的字体格式
        errorDataStyle.setFont(font);
        return errorDataStyle;
    }
}
