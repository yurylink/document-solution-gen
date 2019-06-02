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
import org.apache.poi.xssf.usermodel.XSSFCreationHelper;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.math.BigDecimal;
import java.util.List;

public class ExcelReportBuilder {

    private static final Logger LOGGER = LoggerFactory.getLogger(ExcelReportBuilder.class);

    private final XSSFWorkbook workbook;
    private final Integer columnsQuantity;
    private XSSFSheet sheet;
    private Integer rowNumber;
    private XSSFRow currentRow;
    private Integer cellNumber;
    private CreationHelper creator;

    private ExcelReportBuilder(Integer columnsQuantity){
        this.columnsQuantity = columnsQuantity;

        this.workbook = new XSSFWorkbook();
        this.sheet = workbook.createSheet("report");

        rowNumber = 0;
        cellNumber = 0;
        creator = workbook.getCreationHelper();
    }

    public static ExcelReportBuilder newInstance(Integer columnsQuantity){
        return new ExcelReportBuilder(columnsQuantity);
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

    public ExcelReportBuilder newRow(){
        rowNumber++;
        currentRow = sheet.createRow(rowNumber);
        cellNumber=0;
        return this;
    }

    public ExcelReportBuilder newRowFillEmptyCell(){
        while (cellNumber.compareTo(columnsQuantity) < 0){
            addEmptyCell();
        }
        rowNumber++;
        currentRow = sheet.createRow(rowNumber);
        cellNumber=0;
        return this;
    }

    public ExcelReportBuilder addTitle(String reportName){
        XSSFRow titleRow = sheet.createRow(getRowNumber());
        XSSFCell titleCell = titleRow.createCell(getCellNumber(), XSSFCell.CELL_TYPE_STRING);
        titleCell.setCellStyle(getTitleStyle());
        titleCell.setCellValue(reportName);

        CellRangeAddress cellTitle = new CellRangeAddress(0,1,0, (columnsQuantity -1));

        RegionUtil.setBorderTop(getTitleStyle().getBorderTop(), cellTitle, sheet, workbook);
        RegionUtil.setBorderLeft(getTitleStyle().getBorderLeft(), cellTitle, sheet, workbook);
        RegionUtil.setBorderRight(getTitleStyle().getBorderRight(), cellTitle, sheet, workbook);
        RegionUtil.setBorderBottom(getTitleStyle().getBorderBottom(), cellTitle, sheet, workbook);
        sheet.addMergedRegion(cellTitle);

        sheet.getRow(1).getCell(columnsQuantity-1).getCellStyle().setBorderRight(XSSFCellStyle.BORDER_MEDIUM);

        return this.newRow().newRow();
    }

    public ExcelReportBuilder addSubTitle(String subtitleContent){
        XSSFRow subTitleRow = sheet.createRow(getRowNumber());
        XSSFCell subTitleCell = subTitleRow.createCell(getCellNumber(), XSSFCell.CELL_TYPE_STRING);
        subTitleCell.setCellStyle(getSubTitleStyle());
        subTitleCell.setCellValue(subtitleContent);

        CellRangeAddress subTitleCellRange = new CellRangeAddress(2,3,0, (columnsQuantity -1));

        RegionUtil.setBorderTop(getSubTitleStyle().getBorderTop(), subTitleCellRange, sheet, workbook);
        RegionUtil.setBorderLeft(getSubTitleStyle().getBorderLeft(), subTitleCellRange, sheet, workbook);
        RegionUtil.setBorderRight(getSubTitleStyle().getBorderRight(), subTitleCellRange, sheet, workbook);
        RegionUtil.setBorderBottom(getSubTitleStyle().getBorderBottom(), subTitleCellRange, sheet, workbook);
        sheet.addMergedRegion(subTitleCellRange);

        sheet.getRow(3).getCell(columnsQuantity-1).getCellStyle().setBorderRight(XSSFCellStyle.BORDER_MEDIUM);

        newRow().newRow();
        addEndSectionBorder();
        return this;
    }

    public ExcelReportBuilder addCellRangeAddress(CellRangeAddress cra){
        sheet.addMergedRegion(cra);
        return this;
    }

    public CellRangeAddress getNormalStyleCellAdress(int firstRow, int lastRow, Integer firstCol, Integer lastCol){
        CellRangeAddress cra = new CellRangeAddress(firstRow,lastRow, firstCol,lastCol);
        RegionUtil.setBorderBottom(XSSFCellStyle.BORDER_MEDIUM, cra, sheet, workbook);
        RegionUtil.setBorderRight(XSSFCellStyle.BORDER_MEDIUM, cra, sheet, workbook);
        RegionUtil.setBorderLeft(XSSFCellStyle.BORDER_MEDIUM, cra, sheet, workbook);
        RegionUtil.setBorderBottom(XSSFCellStyle.BORDER_MEDIUM, cra, sheet, workbook);
        return cra;
    }

    public CellRangeAddress getTotalStyleCellAdress(int firstRow, int lastRow, Integer firstCol, Integer lastCol){
        CellRangeAddress cra = new CellRangeAddress(firstRow,lastRow, firstCol,lastCol);
        RegionUtil.setBorderBottom(getLineSeparatorStyle().getBorderBottom(), cra, sheet, workbook);
        RegionUtil.setBorderRight(getLineSeparatorStyle().getBorderRight(), cra, sheet, workbook);
        RegionUtil.setBorderLeft(getLineSeparatorStyle().getBorderLeft(), cra, sheet, workbook);
        RegionUtil.setBorderBottom(getLineSeparatorStyle().getBorderBottom(), cra, sheet, workbook);
        return cra;
    }

    public ExcelReportBuilder addEndSectionBorder(){
        for (int i = 0; i < columnsQuantity; i++) {
            XSSFCell cell = currentRow.createCell(getCellNumber(), XSSFCell.CELL_TYPE_STRING);
            XSSFCellStyle lineSeparatorStyle = workbook.createCellStyle();
            lineSeparatorStyle.setBorderTop(XSSFCellStyle.BORDER_MEDIUM);
            cell.setCellStyle(lineSeparatorStyle);
        }
        return this.newRow();
    }

    public ExcelReportBuilder addHeaderColumns(List<String> headerInfos){
        for (String headerText:headerInfos) {
            addHeaderCell(headerText);
        }
        return this;
    }

    public ExcelReportBuilder addLineSeparatorSection(String content){
        XSSFRow row = sheet.createRow(getRowNumber());
        XSSFCell cell = row.createCell(getCellNumber(), XSSFCell.CELL_TYPE_STRING);
        cell.setCellStyle(getLineSeparatorStyle());
        cell.setCellValue(content);

        CellRangeAddress cellTitle = new CellRangeAddress(
                getRowNumber(),
                getRowNumber(),
                0,
                (columnsQuantity -1));

        sheet.addMergedRegion(cellTitle);

        return this;
    }

    public ExcelReportBuilder addHeaderCell(String content) {
        XSSFCell titleCell = currentRow.createCell(getCellNumber(), XSSFCell.CELL_TYPE_STRING);
        titleCell.setCellValue(content);
        titleCell.setCellStyle(getHeaderStyle());
        return this;
    }

    public ExcelReportBuilder addEmptyCell() {
        XSSFCell cell = currentRow.createCell(getCellNumber(), XSSFCell.CELL_TYPE_BLANK);
        cell.setCellStyle(getNormalStyle());
        return this;
    }

    public ExcelReportBuilder addEmptyTotalCell() {
        XSSFCell cell = currentRow.createCell(getCellNumber(), XSSFCell.CELL_TYPE_BLANK);
        cell.setCellStyle(getLineSeparatorStyle());
        return this;
    }

    public ExcelReportBuilder addTotalCell(BigDecimal value) {
        XSSFCell cell = currentRow.createCell(getCellNumber(), XSSFCell.CELL_TYPE_NUMERIC);
        cell.setCellStyle(getLineSeparatorStyle());

        XSSFCreationHelper chCurrency = workbook.getCreationHelper();
        cell.getCellStyle().setDataFormat(chCurrency.createDataFormat().getFormat("#,##0.00;\\-#,##0.00"));

        cell.setCellValue(value.doubleValue());
        return this;
    }

    public ExcelReportBuilder addTotalCell(String content) {
        XSSFCell cell = currentRow.createCell(getCellNumber(), XSSFCell.CELL_TYPE_STRING);
        cell.setCellStyle(getLineSeparatorStyle());
        cell.setCellValue(creator.createRichTextString(content));
        return this;
    }

    public ExcelReportBuilder addCell(BigDecimal value){
        XSSFCell cell = currentRow.createCell(getCellNumber(), XSSFCell.CELL_TYPE_NUMERIC);
        cell.setCellStyle(getNormalStyle());

        XSSFCreationHelper chCurrency = workbook.getCreationHelper();
        cell.getCellStyle().setDataFormat(chCurrency.createDataFormat().getFormat("#,##0.00;\\-#,##0.00"));

        cell.setCellValue(value.doubleValue());
        return this;
    }

    public ExcelReportBuilder addCell(String content){
        XSSFCell cell = currentRow.createCell(getCellNumber(), XSSFCell.CELL_TYPE_STRING);
        cell.setCellStyle(getNormalStyle());
        cell.setCellValue(creator.createRichTextString(content));
        return this;
    }

    public ExcelReportBuilder addBoldCell(BigDecimal value){
        XSSFCell cell = currentRow.createCell(getCellNumber(), XSSFCell.CELL_TYPE_NUMERIC);
        cell.setCellStyle(getBoldNormalStyle());

        XSSFCreationHelper chCurrency = workbook.getCreationHelper();
        cell.getCellStyle().setDataFormat(chCurrency.createDataFormat().getFormat("#,##0.00;\\-#,##0.00"));

        cell.setCellValue(value.doubleValue());
        return this;
    }

    public ExcelReportBuilder addBoldCell(String content){
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
            sheet.autoSizeColumn(i);
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