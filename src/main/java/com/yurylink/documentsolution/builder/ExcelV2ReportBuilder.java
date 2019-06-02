package com.yurylink.documentsolution.builder;

import com.yurylink.documentsolution.Utils.ExcellReportConstants;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFCreationHelper;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.math.BigDecimal;

public class ExcelV2ReportBuilder {

    private final XSSFWorkbook workbook;
    private final Integer columnsQuantity;
    private XSSFSheet sheet;
    private Integer rowNumber;
    private XSSFRow currentRow;
    private Integer cellNumber;
    private CreationHelper creator;

    private ExcelV2ReportBuilder(Integer columnsQuantity){
        this.columnsQuantity = columnsQuantity;

        this.workbook = new XSSFWorkbook();
        this.sheet = workbook.createSheet("report");

        rowNumber = 0;
        cellNumber = 0;
        this.newRow();
        creator = workbook.getCreationHelper();
    }

    public static ExcelV2ReportBuilder newInstance(Integer columnsQuantity){
        return new ExcelV2ReportBuilder(columnsQuantity);
    }

    public int getRowNumber() {
        return rowNumber;
    }

    public int getCellNumber() {
        Integer currentCellValue = cellNumber;
        cellNumber++;
        return currentCellValue;
    }

    public Integer getCurrentRowNumber() {
        return currentRow.getRowNum();
    }

    public ExcelV2ReportBuilder newRow(){
        rowNumber++;
        currentRow = sheet.createRow(rowNumber);
        cellNumber=0;
        return this;
    }

    public ExcelV2ReportBuilder endOfReportLine(){
        while (cellNumber.compareTo(columnsQuantity) <= 0){
            XSSFCellStyle style = getBorderLessWitheBackgroundStyle();
            style.setBorderBottom(XSSFCellStyle.BORDER_MEDIUM);

            if(cellNumber.equals(0)){
                style.setBorderLeft(XSSFCellStyle.BORDER_MEDIUM);
            }else if (cellNumber.equals(columnsQuantity)) {
                style.setBorderRight(XSSFCellStyle.BORDER_MEDIUM);
            }

            addEmptyCell(style);
        }
        rowNumber++;
        currentRow = sheet.createRow(rowNumber);
        cellNumber=0;
        return this;
    }

    public ExcelV2ReportBuilder newRowFillEmptyCell(){
        while (cellNumber.compareTo(columnsQuantity) <= 0){
            if(cellNumber==columnsQuantity) {
                addEmptyCell(getBorderRightWitheBackgroundStyle(true));
            }
            addEmptyCell(getBorderLessWitheBackgroundStyle());
        }
        rowNumber++;
        currentRow = sheet.createRow(rowNumber);
        cellNumber=0;
        return this;
    }

    public ExcelV2ReportBuilder newRowFillEmptyCell(XSSFCellStyle style){
        while (cellNumber.compareTo(columnsQuantity) < 0){
            addEmptyCell(style);
        }
        rowNumber++;
        currentRow = sheet.createRow(rowNumber);
        cellNumber=0;
        return this;
    }

    public ExcelV2ReportBuilder newRowFillEmptyCellWithContainerMargin(XSSFColor color){
        while (cellNumber.compareTo(columnsQuantity) <= 0){
            XSSFCellStyle style = getStyleRelativeBorder(null);
            style.setFillBackgroundColor(color);
            addEmptyCell(style);
        }
        rowNumber++;
        currentRow = sheet.createRow(rowNumber);
        cellNumber=0;
        return this;
    }

    public ExcelV2ReportBuilder newRowFillEmptyCellWithContainerMargin(){
        while (cellNumber.compareTo(columnsQuantity) <= 0){
            addEmptyCell(getStyleRelativeBorder(null));
        }
        rowNumber++;
        currentRow = sheet.createRow(rowNumber);
        cellNumber=0;
        return this;
    }

    public ExcelV2ReportBuilder newRowFillEmptyCellWithContainerMargin(Integer lvl){
        return this.newRowFillEmptyCellWithContainerMargin(lvl, null);
    }

    public ExcelV2ReportBuilder newRowFillEmptyCellWithContainerMargin(Integer lvl, XSSFColor color){
        while (cellNumber.compareTo(columnsQuantity) <= 0){
            if(color == null){
                addEmptyCell(getStyleRelativeBorder(lvl));
            }else {
                XSSFCellStyle style = getStyleRelativeBorder(lvl);
                style.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
                style.setFillForegroundColor(color);
                addEmptyCell(style);
            }
        }
        rowNumber++;
        currentRow = sheet.createRow(rowNumber);
        cellNumber=0;
        return this;
    }

    private XSSFCellStyle getStyleRelativeBorder(Integer lvl){
        if(cellNumber.equals(0)){
            return getBorderLeftWitheBackgroundStyle(true);
        }else if(cellNumber.equals(columnsQuantity)){
            return getBorderRightWitheBackgroundStyle(true);
        }

        if(lvl != null){
            if(lvl.equals(1)){
                if(cellNumber.equals(1))
                    return getBorderLeftWitheBackgroundStyle();
                else if(cellNumber.equals(columnsQuantity -1))
                    return getBorderRightWitheBackgroundStyle();
            }

            if(lvl.equals(2)){
                if(cellNumber.equals(1) ||
                   cellNumber.equals(2))
                    return getBorderLeftWitheBackgroundStyle();

                else if(cellNumber.equals(columnsQuantity -1) ||
                         cellNumber.equals(columnsQuantity -2))
                    return getBorderRightWitheBackgroundStyle();
            }

            if (lvl.equals(3)){
                if(cellNumber.equals(1) ||
                   cellNumber.equals(2) ||
                   cellNumber.equals(3))
                    return getBorderLeftWitheBackgroundStyle();

                else if(cellNumber.equals(columnsQuantity -1) ||
                         cellNumber.equals(columnsQuantity -2) ||
                         cellNumber.equals(columnsQuantity -3))
                    return getBorderRightWitheBackgroundStyle();
            }
        }
        return getBorderLessWitheBackgroundStyle();
    }

    public XSSFCellStyle getBorderLessWitheBackgroundStyle(){
        XSSFCellStyle style = workbook.createCellStyle();
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setAlignment(HorizontalAlignment.CENTER);

        XSSFFont font = workbook.createFont();
        font.setFontName("SERIF");
        font.setBold(true);
        font.setFontHeight(5);

        style.setFont(font);
        style.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
        style.setFillForegroundColor(ExcellReportConstants.WHITE_COLOR);
        return style;
    }

    private XSSFCellStyle getBorderLeftWitheBackgroundStyle(){
        return getBorderLeftWitheBackgroundStyle(false);
    }

    private XSSFCellStyle getBorderLeftWitheBackgroundStyle(Boolean hasThickerBorders){
        XSSFCellStyle style = workbook.createCellStyle();
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setBorderLeft(hasThickerBorders ? XSSFCellStyle.BORDER_MEDIUM : XSSFCellStyle.BORDER_THIN);

        XSSFFont font = workbook.createFont();
        font.setFontName("SERIF");
        font.setBold(true);
        font.setFontHeight(5);

        style.setFont(font);
        style.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
        style.setFillForegroundColor(ExcellReportConstants.WHITE_COLOR);
        return style;
    }

    private XSSFCellStyle getBorderRightWitheBackgroundStyle(){
        return getBorderRightWitheBackgroundStyle(false);
    }

    private XSSFCellStyle getBorderRightWitheBackgroundStyle(Boolean hasThickerBorder){
        XSSFCellStyle style = workbook.createCellStyle();
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setBorderRight(hasThickerBorder ? XSSFCellStyle.BORDER_MEDIUM : XSSFCellStyle.BORDER_THIN);

        XSSFFont font = workbook.createFont();
        font.setFontName("SERIF");
        font.setBold(true);
        font.setFontHeight(5);

        style.setFont(font);
        style.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
        style.setFillForegroundColor(ExcellReportConstants.WHITE_COLOR);
        return style;
    }

    public ExcelV2ReportBuilder addCellRangeAddress(CellRangeAddress cra){
        sheet.addMergedRegion(cra);
        return this;
    }

    public ExcelV2ReportBuilder addHeaderCell(String content) {
        XSSFCell titleCell = currentRow.createCell(getCellNumber(), XSSFCell.CELL_TYPE_STRING);
        titleCell.setCellValue(content);
        titleCell.setCellStyle(getHeaderStyle());
        return this;
    }

    public ExcelV2ReportBuilder addEmptyCell() {
        XSSFCell cell = currentRow.createCell(getCellNumber(), XSSFCell.CELL_TYPE_BLANK);
        cell.setCellStyle(getNormalStyle());
        return this;
    }

    public ExcelV2ReportBuilder addEmptyCell(Integer lvl, XSSFColor color) {
        XSSFCellStyle style = getStyleRelativeBorder(lvl);
        style.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
        style.setFillForegroundColor(color);
        XSSFCell cell = currentRow.createCell(getCellNumber(), XSSFCell.CELL_TYPE_BLANK);
        cell.setCellStyle(style);
        return this;
    }

    public ExcelV2ReportBuilder addEmptyCell(Integer lvl) {
        XSSFCellStyle style = getStyleRelativeBorder(lvl);
        XSSFCell cell = currentRow.createCell(getCellNumber(), XSSFCell.CELL_TYPE_BLANK);
        cell.setCellStyle(style);
        return this;
    }

    public ExcelV2ReportBuilder addEmptyCell(XSSFCellStyle style) {
        XSSFCell cell = currentRow.createCell(getCellNumber(), XSSFCell.CELL_TYPE_BLANK);
        cell.setCellStyle(style);
        return this;
    }

    public ExcelV2ReportBuilder addEmptyCellWithBorder() {
        if(cellNumber.equals(0) ||
                cellNumber.equals(1) ||
                cellNumber.equals(2)){

            addEmptyCell(getBorderLeftWitheBackgroundStyle());

        }else if(cellNumber.equals(columnsQuantity) ||
                cellNumber.equals(columnsQuantity -1) ||
                cellNumber.equals(columnsQuantity -2)){

            addEmptyCell(getBorderRightWitheBackgroundStyle());
        }else{
            addEmptyCell(getBorderLessWitheBackgroundStyle());
        }
        return this;
    }

    public ExcelV2ReportBuilder addReportHeader(String text){
        return addReportHeader(text, false, false);
    }

    public ExcelV2ReportBuilder addReportHeader(String text, Boolean isTop, Boolean isBottom){
        XSSFCell cell = currentRow.createCell(getCellNumber(), XSSFCell.CELL_TYPE_STRING);
        cell.setCellValue(text);

        XSSFCellStyle style = workbook.createCellStyle();
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setBorderTop(getBorderVariable(isTop));
        style.setBorderLeft(XSSFCellStyle.BORDER_MEDIUM);
        style.setBorderRight(XSSFCellStyle.BORDER_MEDIUM);
        style.setBorderBottom(getBorderVariable(isBottom));

        XSSFFont font = workbook.createFont();
        font.setFontName("SERIF");
        font.setBold(true);
        font.setFontHeight(12);

        style.setFont(font);
        style.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
        style.setFillForegroundColor(ExcellReportConstants.WHITE_COLOR);

        cell.setCellStyle(style);

        CellRangeAddress range = new CellRangeAddress(rowNumber, rowNumber, 0, columnsQuantity);

        RegionUtil.setBorderTop(getBorderVariable(isTop), range, sheet, workbook);
        RegionUtil.setBorderLeft(XSSFCellStyle.BORDER_MEDIUM, range, sheet, workbook);
        RegionUtil.setBorderRight(XSSFCellStyle.BORDER_MEDIUM, range, sheet, workbook);
        RegionUtil.setBorderBottom(getBorderVariable(isBottom), range, sheet, workbook);

        sheet.addMergedRegion(range);
        newRowFillEmptyCell(style);
        return this;
    }

    public ExcelV2ReportBuilder addT1ContainersHeader(String text, XSSFColor backGroundColor){
        addEmptyCellWithBorder();

        XSSFCell cell = currentRow.createCell(getCellNumber(), XSSFCell.CELL_TYPE_STRING);
        cell.setCellValue(text);

        XSSFCellStyle style = workbook.createCellStyle();
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setAlignment(HorizontalAlignment.LEFT);
        style.setBorderTop(XSSFCellStyle.BORDER_THIN);
        style.setBorderLeft(XSSFCellStyle.BORDER_THIN);
        style.setBorderRight(XSSFCellStyle.BORDER_THIN);
        style.setBorderBottom(XSSFCellStyle.BORDER_THIN);

        XSSFFont font = workbook.createFont();
        font.setFontName("SERIF");
        font.setBold(true);
        font.setFontHeight(12);

        style.setFont(font);
        style.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
        style.setFillForegroundColor(backGroundColor);//PAREI AQUI

        cell.setCellStyle(style);

        CellRangeAddress range = new CellRangeAddress(rowNumber, rowNumber, 1, (columnsQuantity-1));
        cellNumber = columnsQuantity;

        RegionUtil.setBorderTop(XSSFCellStyle.BORDER_THIN, range, sheet, workbook);
        RegionUtil.setBorderLeft(XSSFCellStyle.BORDER_THIN, range, sheet, workbook);
        RegionUtil.setBorderRight(XSSFCellStyle.BORDER_THIN, range, sheet, workbook);
        RegionUtil.setBorderBottom(XSSFCellStyle.BORDER_THIN, range, sheet, workbook);

        sheet.addMergedRegion(range);
        newRowFillEmptyCellWithContainerMargin(1);
        return this;
    }

    public ExcelV2ReportBuilder addT1ContainersTrailer(String text, BigDecimal value, XSSFColor backGroundColor){
        addEmptyCellWithBorder();

        XSSFCell cell = currentRow.createCell(getCellNumber(), XSSFCell.CELL_TYPE_STRING);
        cell.setCellValue(text);

        XSSFCellStyle style = workbook.createCellStyle();
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setAlignment(HorizontalAlignment.LEFT);
        style.setBorderTop(XSSFCellStyle.BORDER_THIN);
        style.setBorderLeft(XSSFCellStyle.BORDER_THIN);
        style.setBorderRight(XSSFCellStyle.BORDER_NONE);
        style.setBorderBottom(XSSFCellStyle.BORDER_THIN);

        XSSFFont font = workbook.createFont();
        font.setFontName("SERIF");
        font.setBold(true);
        font.setFontHeight(12);

        style.setFont(font);
        style.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
        style.setFillForegroundColor(backGroundColor);

        cell.setCellStyle(style);

        CellRangeAddress range = new CellRangeAddress(rowNumber, rowNumber, 1, (columnsQuantity-2));

        RegionUtil.setBorderTop(XSSFCellStyle.BORDER_THIN, range, sheet, workbook);
        RegionUtil.setBorderLeft(XSSFCellStyle.BORDER_THIN, range, sheet, workbook);
        RegionUtil.setBorderRight(XSSFCellStyle.BORDER_NONE, range, sheet, workbook);
        RegionUtil.setBorderBottom(XSSFCellStyle.BORDER_THIN, range, sheet, workbook);

        sheet.addMergedRegion(range);

        //--------------NUMERIC VALUE------------------------------------------------------------//
        XSSFCreationHelper chCurrency = workbook.getCreationHelper();
        XSSFCell cell2 = currentRow.createCell((columnsQuantity-1), XSSFCell.CELL_TYPE_NUMERIC);
        cellNumber = columnsQuantity;

        XSSFCellStyle style2 = workbook.createCellStyle();

        style2.setAlignment(CellStyle.ALIGN_CENTER);
        style2.setVerticalAlignment(VerticalAlignment.CENTER);
        style2.setAlignment(HorizontalAlignment.RIGHT);
        style2.setBorderTop(XSSFCellStyle.BORDER_THIN);
        style2.setBorderLeft(XSSFCellStyle.BORDER_NONE);
        style2.setBorderRight(XSSFCellStyle.BORDER_THIN);
        style2.setBorderBottom(XSSFCellStyle.BORDER_THIN);
        style2.setDataFormat(chCurrency.createDataFormat().getFormat("#,##0.00;\\-#,##0.00"));

        style2.setFont(font);
        style2.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
        style2.setFillForegroundColor(backGroundColor);

        cell2.setCellStyle(style2);
        cell2.setCellValue(value.doubleValue());

        newRowFillEmptyCellWithContainerMargin(1);
        return this;
    }

    public ExcelV2ReportBuilder addT2ContainersHeaderLines(XSSFColor backGroundColor){
        while (cellNumber != columnsQuantity){
            XSSFCellStyle style = getStyleRelativeBorder(2);
            style.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
            style.setFillBackgroundColor(backGroundColor);
            if(cellNumber.compareTo(2)>=0 &&
               cellNumber.compareTo(columnsQuantity-2)<=0) {
                style.setBorderTop(CellStyle.BORDER_THIN);
            }

            addEmptyCell(style);
        }
        newRowFillEmptyCellWithContainerMargin(3);
        return this;
    }

    public ExcelV2ReportBuilder addT3ContainersHeaderLines(XSSFColor backGroundColor){
        while (cellNumber != columnsQuantity){
            XSSFCellStyle style = getStyleRelativeBorder(3);
            if(cellNumber.compareTo(3)>=0 &&
                    cellNumber.compareTo(columnsQuantity-3)<=0)
                style.setBorderBottom(CellStyle.BORDER_THIN);

            addEmptyCell(style);
        }
        newRowFillEmptyCellWithContainerMargin(3);
        return this;
    }

    public ExcelV2ReportBuilder addT2ContainersHeader(String text, XSSFColor backGroundColor){
        addEmptyCellWithBorder();
        addEmptyCellWithBorder();

        XSSFCell cell = currentRow.createCell(getCellNumber(), XSSFCell.CELL_TYPE_STRING);
        cell.setCellValue(text);

        XSSFCellStyle style = workbook.createCellStyle();
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setAlignment(HorizontalAlignment.LEFT);
        style.setBorderTop(XSSFCellStyle.BORDER_THIN);
        style.setBorderLeft(XSSFCellStyle.BORDER_THIN);
        style.setBorderRight(XSSFCellStyle.BORDER_THIN);
        style.setBorderBottom(XSSFCellStyle.BORDER_THIN);

        XSSFFont font = workbook.createFont();
        font.setFontName("SERIF");
        font.setBold(true);
        font.setFontHeight(12);

        style.setFont(font);
        style.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
        style.setFillForegroundColor(backGroundColor);

        cell.setCellStyle(style);

        CellRangeAddress range = new CellRangeAddress(rowNumber, rowNumber, 2, (columnsQuantity-2));
        cellNumber = columnsQuantity-1;

        RegionUtil.setBorderTop(XSSFCellStyle.BORDER_THIN, range, sheet, workbook);
        RegionUtil.setBorderLeft(XSSFCellStyle.BORDER_THIN, range, sheet, workbook);
        RegionUtil.setBorderRight(XSSFCellStyle.BORDER_THIN, range, sheet, workbook);
        RegionUtil.setBorderBottom(XSSFCellStyle.BORDER_THIN, range, sheet, workbook);

        sheet.addMergedRegion(range);
        newRowFillEmptyCellWithContainerMargin(2);
        return this;
    }

    public ExcelV2ReportBuilder addT2ContainersTrailer(String text, BigDecimal value, XSSFColor color){
        addEmptyCellWithBorder();
        addEmptyCellWithBorder();

        XSSFCell cell = currentRow.createCell(getCellNumber(), XSSFCell.CELL_TYPE_STRING);
        cell.setCellValue(text);

        XSSFCellStyle style = workbook.createCellStyle();
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setAlignment(HorizontalAlignment.LEFT);
        style.setBorderTop(XSSFCellStyle.BORDER_THIN);
        style.setBorderLeft(XSSFCellStyle.BORDER_THIN);
        style.setBorderRight(XSSFCellStyle.BORDER_NONE);
        style.setBorderBottom(XSSFCellStyle.BORDER_THIN);

        XSSFFont font = workbook.createFont();
        font.setFontName("SERIF");
        font.setBold(true);
        font.setFontHeight(12);

        style.setFont(font);
        style.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
        style.setFillForegroundColor(color);

        cell.setCellStyle(style);

        CellRangeAddress range = new CellRangeAddress(rowNumber, rowNumber, 2, (columnsQuantity-3));
        cellNumber = columnsQuantity - 2;

        RegionUtil.setBorderTop(XSSFCellStyle.BORDER_THIN, range, sheet, workbook);
        RegionUtil.setBorderLeft(XSSFCellStyle.BORDER_THIN, range, sheet, workbook);
        RegionUtil.setBorderRight(XSSFCellStyle.BORDER_NONE, range, sheet, workbook);
        RegionUtil.setBorderBottom(XSSFCellStyle.BORDER_THIN, range, sheet, workbook);

        sheet.addMergedRegion(range);

        //--------------NUMERIC VALUE------------------------------------------------------------//
        XSSFCreationHelper chCurrency = workbook.getCreationHelper();
        XSSFCell cell2 = currentRow.createCell(getCellNumber(), XSSFCell.CELL_TYPE_NUMERIC);

        XSSFCellStyle style2 = workbook.createCellStyle();

        style2.setAlignment(CellStyle.ALIGN_CENTER);
        style2.setVerticalAlignment(VerticalAlignment.CENTER);
        style2.setAlignment(HorizontalAlignment.RIGHT);
        style2.setBorderTop(XSSFCellStyle.BORDER_THIN);
        style2.setBorderLeft(XSSFCellStyle.BORDER_NONE);
        style2.setBorderRight(XSSFCellStyle.BORDER_THIN);
        style2.setBorderBottom(XSSFCellStyle.BORDER_THIN);
        style2.setDataFormat(chCurrency.createDataFormat().getFormat("#,##0.00;\\-#,##0.00"));

        style2.setFont(font);
        style2.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
        style2.setFillForegroundColor(color);

        cell2.setCellStyle(style2);
        cell2.setCellValue(value.doubleValue());

        newRowFillEmptyCellWithContainerMargin(2);
        return this;
    }

    public ExcelV2ReportBuilder addT2ContainersCell(String text, XSSFColor color){
        XSSFCell cell = currentRow.createCell(getCellNumber(), XSSFCell.CELL_TYPE_STRING);
        cell.setCellValue(text);

        XSSFCellStyle style = workbook.createCellStyle();
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setBorderTop(XSSFCellStyle.BORDER_THIN);
        style.setBorderLeft(XSSFCellStyle.BORDER_THIN);
        style.setBorderRight(XSSFCellStyle.BORDER_THIN);
        style.setBorderBottom(XSSFCellStyle.BORDER_THIN);

        XSSFFont font = workbook.createFont();
        font.setFontName("SERIF");
        font.setBold(true);
        font.setFontHeight(10);

        style.setFont(font);
        style.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
        style.setFillForegroundColor(color);

        cell.setCellStyle(style);
        return this;
    }

    public ExcelV2ReportBuilder addT2ContainersCell(BigDecimal value, XSSFColor color){
        XSSFFont font = workbook.createFont();
        font.setFontName("SERIF");
        font.setBold(true);
        font.setFontHeight(10);

        //--------------NUMERIC VALUE------------------------------------------------------------//
        XSSFCreationHelper chCurrency = workbook.getCreationHelper();
        XSSFCell cell2 = currentRow.createCell(getCellNumber(), XSSFCell.CELL_TYPE_NUMERIC);

        XSSFCellStyle style2 = workbook.createCellStyle();

        style2.setAlignment(CellStyle.ALIGN_CENTER);
        style2.setVerticalAlignment(VerticalAlignment.CENTER);
        style2.setAlignment(HorizontalAlignment.CENTER);
        style2.setBorderTop(XSSFCellStyle.BORDER_THIN);
        style2.setBorderLeft(XSSFCellStyle.BORDER_NONE);
        style2.setBorderRight(XSSFCellStyle.BORDER_THIN);
        style2.setBorderBottom(XSSFCellStyle.BORDER_THIN);
        style2.setDataFormat(chCurrency.createDataFormat().getFormat("#,##0.00;\\-#,##0.00"));

        style2.setFont(font);
        style2.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
        style2.setFillForegroundColor(color);

        cell2.setCellStyle(style2);
        cell2.setCellValue(value.doubleValue());
        return this;
    }

    public ExcelV2ReportBuilder addT3TextHeader(String text){
        XSSFCell cell = currentRow.createCell(getCellNumber(), XSSFCell.CELL_TYPE_STRING);
        cell.setCellValue(text);

        XSSFCellStyle style = workbook.createCellStyle();
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setAlignment(HorizontalAlignment.LEFT);

        XSSFFont font = workbook.createFont();
        font.setFontName("SERIF");
        font.setBold(true);
        font.setFontHeight(10);

        style.setFont(font);
        style.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
        style.setFillForegroundColor(ExcellReportConstants.WHITE_COLOR);

        cell.setCellStyle(style);
        return this;
    }

    public ExcelV2ReportBuilder addT3ContainersHeader(String text, XSSFColor backGroundColor){
        XSSFCell cell = currentRow.createCell(getCellNumber(), XSSFCell.CELL_TYPE_STRING);
        cell.setCellValue(text);

        XSSFCellStyle style = workbook.createCellStyle();
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setAlignment(HorizontalAlignment.LEFT);
        style.setBorderTop(XSSFCellStyle.BORDER_THIN);
        style.setBorderLeft(XSSFCellStyle.BORDER_THIN);
        style.setBorderRight(XSSFCellStyle.BORDER_THIN);
        style.setBorderBottom(XSSFCellStyle.BORDER_THIN);

        XSSFFont font = workbook.createFont();
        font.setFontName("SERIF");
        font.setBold(true);
        font.setFontHeight(10);

        style.setFont(font);
        style.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
        style.setFillForegroundColor(backGroundColor);

        cell.setCellStyle(style);
        return this;
    }

    public ExcelV2ReportBuilder addT3ContainersTrailer(String text, BigDecimal value, XSSFColor color){
        addEmptyCellWithBorder();
        addEmptyCellWithBorder();
        addEmptyCellWithBorder();

        XSSFCell cell = currentRow.createCell(getCellNumber(), XSSFCell.CELL_TYPE_STRING);
        cell.setCellValue(text);

        XSSFCellStyle style = workbook.createCellStyle();
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setAlignment(HorizontalAlignment.LEFT);
        style.setBorderTop(XSSFCellStyle.BORDER_THIN);
        style.setBorderLeft(XSSFCellStyle.BORDER_THIN);
        style.setBorderRight(XSSFCellStyle.BORDER_NONE);
        style.setBorderBottom(XSSFCellStyle.BORDER_THIN);

        XSSFFont font = workbook.createFont();
        font.setFontName("SERIF");
        font.setBold(true);
        font.setFontHeight(12);

        style.setFont(font);
        style.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
        style.setFillForegroundColor(color);

        cell.setCellStyle(style);

        CellRangeAddress range = new CellRangeAddress(rowNumber, rowNumber, 3, (columnsQuantity-4));
        cellNumber = columnsQuantity - 3;

        RegionUtil.setBorderTop(XSSFCellStyle.BORDER_THIN, range, sheet, workbook);
        RegionUtil.setBorderLeft(XSSFCellStyle.BORDER_THIN, range, sheet, workbook);
        RegionUtil.setBorderRight(XSSFCellStyle.BORDER_NONE, range, sheet, workbook);
        RegionUtil.setBorderBottom(XSSFCellStyle.BORDER_THIN, range, sheet, workbook);

        sheet.addMergedRegion(range);

        //--------------NUMERIC VALUE------------------------------------------------------------//
        XSSFCreationHelper chCurrency = workbook.getCreationHelper();
        XSSFCell cell2 = currentRow.createCell(getCellNumber(), XSSFCell.CELL_TYPE_NUMERIC);

        XSSFCellStyle style2 = workbook.createCellStyle();

        style2.setAlignment(CellStyle.ALIGN_CENTER);
        style2.setVerticalAlignment(VerticalAlignment.CENTER);
        style2.setAlignment(HorizontalAlignment.RIGHT);
        style2.setBorderTop(XSSFCellStyle.BORDER_THIN);
        style2.setBorderLeft(XSSFCellStyle.BORDER_NONE);
        style2.setBorderRight(XSSFCellStyle.BORDER_THIN);
        style2.setBorderBottom(XSSFCellStyle.BORDER_THIN);
        style2.setDataFormat(chCurrency.createDataFormat().getFormat("#,##0.00;\\-#,##0.00"));

        style2.setFont(font);
        style2.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
        style2.setFillForegroundColor(ExcellReportConstants.SOFT_GREY);

        cell2.setCellStyle(style2);
        cell2.setCellValue(value.doubleValue());

        newRowFillEmptyCellWithContainerMargin(3);
        return this;
    }

    private short getBorderVariable(Boolean hasBorder){
        return hasBorder ? XSSFCellStyle.BORDER_MEDIUM : XSSFCellStyle.BORDER_NONE;
    }

    public ExcelV2ReportBuilder addEmptyTotalCell() {
        XSSFCell cell = currentRow.createCell(getCellNumber(), XSSFCell.CELL_TYPE_BLANK);
        cell.setCellStyle(getLineSeparatorStyle());
        return this;
    }

    public ExcelV2ReportBuilder addTotalCell(BigDecimal value) {
        XSSFCell cell = currentRow.createCell(getCellNumber(), XSSFCell.CELL_TYPE_NUMERIC);
        cell.setCellStyle(getLineSeparatorStyle());

        XSSFCreationHelper chCurrency = workbook.getCreationHelper();
        cell.getCellStyle().setDataFormat(chCurrency.createDataFormat().getFormat("#,##0.00;\\-#,##0.00"));

        cell.setCellValue(value.doubleValue());
        return this;
    }

    public ExcelV2ReportBuilder addTotalCell(String content) {
        XSSFCell cell = currentRow.createCell(getCellNumber(), XSSFCell.CELL_TYPE_STRING);
        cell.setCellStyle(getLineSeparatorStyle());
        cell.setCellValue(creator.createRichTextString(content));
        return this;
    }

    public ExcelV2ReportBuilder addCell(BigDecimal value, XSSFColor color){
        XSSFCellStyle style = workbook.createCellStyle();
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setBorderTop(XSSFCellStyle.BORDER_THIN);
        style.setBorderLeft(XSSFCellStyle.BORDER_THIN);
        style.setBorderRight(XSSFCellStyle.BORDER_THIN);
        style.setBorderBottom(XSSFCellStyle.BORDER_THIN);

        XSSFFont font = workbook.createFont();
        font.setFontName("SERIF");
        font.setBold(false);
        font.setFontHeight(10);

        style.setFont(font);

        XSSFCell cell = currentRow.createCell(getCellNumber(), XSSFCell.CELL_TYPE_NUMERIC);
//        style.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
        style.setFillBackgroundColor(color);

        cell.setCellStyle(style);

        XSSFCreationHelper chCurrency = workbook.getCreationHelper();
        cell.getCellStyle().setDataFormat(chCurrency.createDataFormat().getFormat("#,##0.00;\\-#,##0.00"));

        cell.setCellValue(value.doubleValue());
        return this;
    }

    public ExcelV2ReportBuilder addCell(BigDecimal value){
        XSSFCell cell = currentRow.createCell(getCellNumber(), XSSFCell.CELL_TYPE_NUMERIC);
        cell.setCellStyle(getNormalStyle());

        XSSFCreationHelper chCurrency = workbook.getCreationHelper();
        cell.getCellStyle().setDataFormat(chCurrency.createDataFormat().getFormat("#,##0.00;\\-#,##0.00"));

        cell.setCellValue(value.doubleValue());
        return this;
    }

    public ExcelV2ReportBuilder addCell(Integer lvl, BigDecimal value){
        XSSFCell cell = currentRow.createCell(getCellNumber(), XSSFCell.CELL_TYPE_NUMERIC);
        XSSFCellStyle style = getStyleRelativeBorder(lvl);
        style.getFont().setFontHeight(10);
        style.getFont().setBold(false);
        cell.setCellStyle(style);

        XSSFCreationHelper chCurrency = workbook.getCreationHelper();
        cell.getCellStyle().setDataFormat(chCurrency.createDataFormat().getFormat("#,##0.00;\\-#,##0.00"));

        cell.setCellValue(value.doubleValue());
        return this;
    }

    public ExcelV2ReportBuilder addCell(Integer lvl, String content){
        XSSFCellStyle style = getStyleRelativeBorder(lvl);
        XSSFCell cell = currentRow.createCell(getCellNumber(), XSSFCell.CELL_TYPE_STRING);

        XSSFFont font = workbook.createFont();
        font.setFontName("SERIF");
        font.setBold(false);
        font.setFontHeight(10);

        style.setFont(font);
        cell.setCellStyle(style);
        cell.setCellValue(creator.createRichTextString(content));
        return this;
    }


    public ExcelV2ReportBuilder addCell(String content, XSSFColor color){
        XSSFCellStyle style = workbook.createCellStyle();
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setBorderTop(XSSFCellStyle.BORDER_THIN);
        style.setBorderLeft(XSSFCellStyle.BORDER_THIN);
        style.setBorderRight(XSSFCellStyle.BORDER_THIN);
        style.setBorderBottom(XSSFCellStyle.BORDER_THIN);

        XSSFFont font = workbook.createFont();
        font.setFontName("SERIF");
        font.setBold(false);
        font.setFontHeight(10);

        style.setFont(font);

        XSSFCell cell = currentRow.createCell(getCellNumber(), XSSFCell.CELL_TYPE_STRING);
//        style.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
        style.setFillBackgroundColor(color);
        cell.setCellStyle(style);
        cell.setCellValue(creator.createRichTextString(content));
        return this;
    }

    public ExcelV2ReportBuilder addCell(String content, XSSFCellStyle style){
        XSSFCell cell = currentRow.createCell(getCellNumber(), XSSFCell.CELL_TYPE_STRING);
        cell.setCellStyle(style);
        cell.setCellValue(creator.createRichTextString(content));
        return this;
    }

    public ExcelV2ReportBuilder addT3Cell(String content){
        XSSFCellStyle style = getStyleRelativeBorder(3);
        XSSFCell cell = currentRow.createCell(getCellNumber(), XSSFCell.CELL_TYPE_STRING);
        style.setBorderRight(style.getBorderTop() == XSSFCellStyle.BORDER_MEDIUM ? XSSFCellStyle.BORDER_MEDIUM : XSSFCellStyle.BORDER_THIN);
        style.setBorderRight(style.getBorderLeft() == XSSFCellStyle.BORDER_MEDIUM ? XSSFCellStyle.BORDER_MEDIUM : XSSFCellStyle.BORDER_THIN);
        style.setBorderRight(style.getBorderRight() == XSSFCellStyle.BORDER_MEDIUM ? XSSFCellStyle.BORDER_MEDIUM : XSSFCellStyle.BORDER_THIN);
        style.setBorderRight(style.getBorderBottom() == XSSFCellStyle.BORDER_MEDIUM ? XSSFCellStyle.BORDER_MEDIUM : XSSFCellStyle.BORDER_THIN);
        style.getFont().setFontHeight(10);
        style.getFont().setBold(false);
        cell.setCellStyle(style);
        cell.setCellValue(creator.createRichTextString(content));
        return this;
    }

    public ExcelV2ReportBuilder addT3Cell(BigDecimal value){
        XSSFCellStyle style = getStyleRelativeBorder(3);
        XSSFCell cell = currentRow.createCell(getCellNumber(), XSSFCell.CELL_TYPE_NUMERIC);

        style.setBorderRight(style.getBorderTop() == XSSFCellStyle.BORDER_MEDIUM ? XSSFCellStyle.BORDER_MEDIUM : XSSFCellStyle.BORDER_THIN);
        style.setBorderRight(style.getBorderLeft() == XSSFCellStyle.BORDER_MEDIUM ? XSSFCellStyle.BORDER_MEDIUM : XSSFCellStyle.BORDER_THIN);
        style.setBorderRight(style.getBorderRight() == XSSFCellStyle.BORDER_MEDIUM ? XSSFCellStyle.BORDER_MEDIUM : XSSFCellStyle.BORDER_THIN);
        style.setBorderRight(style.getBorderBottom() == XSSFCellStyle.BORDER_MEDIUM ? XSSFCellStyle.BORDER_MEDIUM : XSSFCellStyle.BORDER_THIN);
        style.getFont().setBold(false);
        style.getFont().setFontHeight(10);
        cell.setCellStyle(style);

        XSSFCreationHelper chCurrency = workbook.getCreationHelper();
        cell.getCellStyle().setDataFormat(chCurrency.createDataFormat().getFormat("#,##0.00;\\-#,##0.00"));

        cell.setCellValue(value.doubleValue());
        return this;
    }

    public ExcelV2ReportBuilder addT3BoldCell(String content){
        XSSFCellStyle style = getStyleRelativeBorder(3);
        XSSFCell cell = currentRow.createCell(getCellNumber(), XSSFCell.CELL_TYPE_STRING);
        style.setBorderRight(style.getBorderTop() == XSSFCellStyle.BORDER_MEDIUM ? XSSFCellStyle.BORDER_MEDIUM : XSSFCellStyle.BORDER_THIN);
        style.setBorderRight(style.getBorderLeft() == XSSFCellStyle.BORDER_MEDIUM ? XSSFCellStyle.BORDER_MEDIUM : XSSFCellStyle.BORDER_THIN);
        style.setBorderRight(style.getBorderRight() == XSSFCellStyle.BORDER_MEDIUM ? XSSFCellStyle.BORDER_MEDIUM : XSSFCellStyle.BORDER_THIN);
        style.setBorderRight(style.getBorderBottom() == XSSFCellStyle.BORDER_MEDIUM ? XSSFCellStyle.BORDER_MEDIUM : XSSFCellStyle.BORDER_THIN);
//        style.getFont().setBold(true);
        style.getFont().setFontHeight(10);
        cell.setCellStyle(style);
        cell.setCellValue(creator.createRichTextString(content));
        return this;
    }

    public ExcelV2ReportBuilder addT3BoldCell(BigDecimal value){
        XSSFCellStyle style = getStyleRelativeBorder(3);
        XSSFCell cell = currentRow.createCell(getCellNumber(), XSSFCell.CELL_TYPE_NUMERIC);

        style.setBorderRight(style.getBorderTop() == XSSFCellStyle.BORDER_MEDIUM ? XSSFCellStyle.BORDER_MEDIUM : XSSFCellStyle.BORDER_THIN);
        style.setBorderRight(style.getBorderLeft() == XSSFCellStyle.BORDER_MEDIUM ? XSSFCellStyle.BORDER_MEDIUM : XSSFCellStyle.BORDER_THIN);
        style.setBorderRight(style.getBorderRight() == XSSFCellStyle.BORDER_MEDIUM ? XSSFCellStyle.BORDER_MEDIUM : XSSFCellStyle.BORDER_THIN);
        style.setBorderRight(style.getBorderBottom() == XSSFCellStyle.BORDER_MEDIUM ? XSSFCellStyle.BORDER_MEDIUM : XSSFCellStyle.BORDER_THIN);
//        style.getFont().setBold(true);
        style.getFont().setFontHeight(10);
        cell.setCellStyle(style);

        XSSFCreationHelper chCurrency = workbook.getCreationHelper();
        cell.getCellStyle().setDataFormat(chCurrency.createDataFormat().getFormat("#,##0.00;\\-#,##0.00"));

        cell.setCellValue(value.doubleValue());
        return this;
    }

    public ExcelV2ReportBuilder addCell(String content){
        XSSFCell cell = currentRow.createCell(getCellNumber(), XSSFCell.CELL_TYPE_STRING);
        cell.setCellStyle(getNormalStyle());
        cell.setCellValue(creator.createRichTextString(content));
        return this;
    }

    public ExcelV2ReportBuilder addBoldCell(BigDecimal value){
        XSSFCell cell = currentRow.createCell(getCellNumber(), XSSFCell.CELL_TYPE_NUMERIC);
        cell.setCellStyle(getBoldNormalStyle());

        XSSFCreationHelper chCurrency = workbook.getCreationHelper();
        cell.getCellStyle().setDataFormat(chCurrency.createDataFormat().getFormat("#,##0.00;\\-#,##0.00"));

        cell.setCellValue(value.doubleValue());
        return this;
    }

    public ExcelV2ReportBuilder addBoldCell(String content){
        XSSFCell cell = currentRow.createCell(getCellNumber(), XSSFCell.CELL_TYPE_STRING);
        cell.setCellStyle(getBoldNormalStyle());
        cell.setCellValue(content);
        return this;
    }

    private XSSFCellStyle getNormalStyle(){
        XSSFCellStyle style = workbook.createCellStyle();
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setBorderTop(XSSFCellStyle.BORDER_THIN);
        style.setBorderLeft(cellNumber.compareTo(0)==0 ? XSSFCellStyle.BORDER_MEDIUM : XSSFCellStyle.BORDER_THIN);
        style.setBorderRight(cellNumber.compareTo(columnsQuantity)==0 ? XSSFCellStyle.BORDER_MEDIUM : XSSFCellStyle.BORDER_THIN);
        style.setBorderBottom(XSSFCellStyle.BORDER_THIN);

        XSSFFont font = workbook.createFont();
        font.setFontName("SERIF");
        font.setBold(false);
        font.setFontHeight(10);

        style.setFont(font);
        return style;
    }

    private XSSFCellStyle getNormalBorderlessStyle(){
        XSSFCellStyle style = workbook.createCellStyle();
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setBorderTop(XSSFCellStyle.BORDER_NONE);
        style.setBorderLeft(XSSFCellStyle.BORDER_NONE);
        style.setBorderRight(XSSFCellStyle.BORDER_NONE);
        style.setBorderBottom(XSSFCellStyle.BORDER_NONE);

        XSSFFont font = workbook.createFont();
        font.setFontName("SERIF");
        font.setBold(false);
        font.setFontHeight(10);

        style.setFont(font);
        return style;
    }

    private XSSFCellStyle getBoldNormalStyle(){
        XSSFCellStyle style = workbook.createCellStyle();
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setBorderTop(XSSFCellStyle.BORDER_THIN);
        style.setBorderLeft(cellNumber.compareTo(0)==0 ? XSSFCellStyle.BORDER_MEDIUM : XSSFCellStyle.BORDER_THIN);
        style.setBorderRight(cellNumber.compareTo(columnsQuantity)==0 ? XSSFCellStyle.BORDER_MEDIUM : XSSFCellStyle.BORDER_THIN);
        style.setBorderBottom(XSSFCellStyle.BORDER_THIN);

        XSSFFont font = workbook.createFont();
        font.setFontName("SERIF");
        font.setBold(true);
        font.setBoldweight(Font.BOLDWEIGHT_BOLD);
        font.setFontHeight(11);

        style.setFont(font);
        return style;
    }

    private XSSFCellStyle getLineSeparatorStyle(){
        XSSFCellStyle style = workbook.createCellStyle();
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setBorderTop(XSSFCellStyle.BORDER_THIN);
        style.setBorderLeft(XSSFCellStyle.BORDER_MEDIUM);
        style.setBorderRight(XSSFCellStyle.BORDER_MEDIUM);
        style.setBorderBottom(XSSFCellStyle.BORDER_THIN);

        XSSFFont font = workbook.createFont();
        font.setFontName("SERIF");
        font.setBold(true);
        font.setBoldweight(Font.BOLDWEIGHT_BOLD);
        font.setFontHeight(12);
        font.setColor(ExcellReportConstants.WHITE_COLOR);

        style.setFont(font);
        style.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
        style.setFillForegroundColor(ExcellReportConstants.DARK_GREY);
        return style;
    }

    private XSSFCellStyle getHeaderStyle(){
        XSSFCellStyle titleCellStyle = workbook.createCellStyle();
        titleCellStyle.setAlignment(CellStyle.ALIGN_CENTER);
        titleCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        titleCellStyle.setAlignment(HorizontalAlignment.CENTER);
        titleCellStyle.setBorderTop(XSSFCellStyle.BORDER_MEDIUM);
        titleCellStyle.setBorderLeft(cellNumber.compareTo(0)==0 ? XSSFCellStyle.BORDER_MEDIUM : XSSFCellStyle.BORDER_MEDIUM);
        titleCellStyle.setBorderRight(cellNumber.compareTo(columnsQuantity)==0 ? XSSFCellStyle.BORDER_MEDIUM : XSSFCellStyle.BORDER_MEDIUM);
        titleCellStyle.setBorderBottom(XSSFCellStyle.BORDER_NONE);

        XSSFFont titleFont = workbook.createFont();
        titleFont.setFontName("SERIF");
        titleFont.setBold(true);
        titleFont.setFontHeight(12);

        titleCellStyle.setFont(titleFont);
        return titleCellStyle;
    }


    private XSSFCellStyle getTitleStyle(){
        XSSFCellStyle titleCellStyle = workbook.createCellStyle();
        titleCellStyle.setAlignment(CellStyle.ALIGN_CENTER);
        titleCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        titleCellStyle.setAlignment(HorizontalAlignment.CENTER);
        titleCellStyle.setBorderTop(XSSFCellStyle.BORDER_MEDIUM);
        titleCellStyle.setBorderLeft(XSSFCellStyle.BORDER_MEDIUM);
        titleCellStyle.setBorderRight(XSSFCellStyle.BORDER_MEDIUM);
        titleCellStyle.setBorderBottom(XSSFCellStyle.BORDER_NONE);

        XSSFFont titleFont = workbook.createFont();
        titleFont.setFontName("SERIF");
        titleFont.setBold(true);
        titleFont.setBoldweight(Font.BOLDWEIGHT_BOLD);
        titleFont.setFontHeight(16);

        titleCellStyle.setFont(titleFont);
        titleCellStyle.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
        titleCellStyle.setFillForegroundColor(ExcellReportConstants.SOFT_GREY);
        return titleCellStyle;
    }

    private XSSFCellStyle getSubTitleStyle(){
        XSSFCellStyle subTitleStyle = workbook.createCellStyle();
        subTitleStyle.setAlignment(CellStyle.ALIGN_CENTER);
        subTitleStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        subTitleStyle.setAlignment(HorizontalAlignment.CENTER);
        subTitleStyle.setBorderTop(XSSFCellStyle.BORDER_NONE);
        subTitleStyle.setBorderLeft(XSSFCellStyle.BORDER_MEDIUM);
        subTitleStyle.setBorderRight(XSSFCellStyle.BORDER_MEDIUM);
        subTitleStyle.setBorderBottom(XSSFCellStyle.BORDER_MEDIUM);

        XSSFFont titleFont = workbook.createFont();
        titleFont.setFontName("SERIF");
        titleFont.setBold(true);
        titleFont.setBoldweight(Font.BOLDWEIGHT_NORMAL);
        titleFont.setFontHeight(13);

        subTitleStyle.setFont(titleFont);
        subTitleStyle.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
        subTitleStyle.setFillForegroundColor(ExcellReportConstants.SOFT_GREY);
        return subTitleStyle;
    }

    public ByteArrayInputStream build() throws FileNotFoundException, IOException  {
        for (int i = 0; i < columnsQuantity; i++) {
            sheet.autoSizeColumn(i);
        }

        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        workbook.write(outputStream);

        return new ByteArrayInputStream(outputStream.toByteArray());
    }

    public void buildFileOutputStream(){
        for (int i = 0; i < columnsQuantity ; i++) {
            sheet.autoSizeColumn(i, true);
        }


    }

//    public void fecharArquivoTeste() throws FileNotFoundException, IOException{
//        for (int i = 0; i < columnsQuantity; i++) {
//            sheet.autoSizeColumn(i);
//        }
//        FileOutputStream file = new FileOutputStream(new File("/home/wesley/4ward/saida/teste.xlsx"));
//        workbook.write(file);
//        file.close();
//    }
}