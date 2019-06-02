package com.yurylink.documentsolution.builder;

import com.itextpdf.text.BaseColor;
import com.itextpdf.text.Font;
import com.itextpdf.text.FontFactory;

public class ExcelUtil {

    public static Font getFont8() {
        return FontFactory.getFont(FontFactory.HELVETICA, 8, Font.NORMAL, new BaseColor(0, 0, 0));
    }
    public static Font getFont(int tamanho, boolean bold) {
        if (bold) {
            return FontFactory.getFont(FontFactory.HELVETICA, tamanho, Font.BOLD, new BaseColor(0, 0, 0));
        } else {
            return FontFactory.getFont(FontFactory.HELVETICA, tamanho, Font.NORMAL, new BaseColor(0, 0, 0));
        }
    }
//
//
//        public static Cell createCell(Row row, int posicao, CellStyle style){
//            Cell cell = row.createCell(posicao);
//            cell.setCellStyle(style);
//            return cell;
//
//        }
//
//        public static Cell createCell(Row row, int posicao, CellStyle style, int colspan){
//            Cell cell = row.createCell(posicao);
//            cell.setCellStyle(style);
//            return cell;
//
//        }
//
//        public static Map<String, CellStyle> createStyles(Workbook wb){
//            Map<String, CellStyle> styles = new HashMap<String, CellStyle>();
//            CellStyle style;
//            org.apache.poi.ss.usermodel.Font titleFont = wb.createFont();
//            titleFont.setFontHeightInPoints((short)18);
//            titleFont.setBoldweight((short) Font.BOLD);
//            style = wb.createCellStyle();
//            style.setAlignment(CellStyle.ALIGN_CENTER);
//            style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
//            style.setFont(titleFont);
//            styles.put("title", style);
//
//            org.apache.poi.ss.usermodel.Font monthFont = wb.createFont();
//            monthFont.setFontHeightInPoints((short)11);
//            titleFont.setBoldweight((short) Font.BOLD);
//            monthFont.setColor(IndexedColors.WHITE.getIndex());
//            style = wb.createCellStyle();
//            style.setAlignment(CellStyle.ALIGN_CENTER);
//            style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
//            style.setFillForegroundColor(IndexedColors.BLACK.getIndex());
//            style.setFillPattern(CellStyle.SOLID_FOREGROUND);
//            style.setFont(monthFont);
//            style.setWrapText(true);
//            styles.put("header", style);
//
//
//            org.apache.poi.ss.usermodel.Font redFont = wb.createFont();
//            redFont.setFontHeightInPoints((short)11);
//            redFont.setColor(IndexedColors.WHITE.getIndex());
//            style = wb.createCellStyle();
//            style.setAlignment(CellStyle.ALIGN_CENTER);
//            style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
//            style.setFillForegroundColor(IndexedColors.RED.getIndex());
//            style.setFillPattern(CellStyle.SOLID_FOREGROUND);
//            style.setFont(redFont);
//            style.setWrapText(true);
//            style.setBorderRight(CellStyle.BORDER_THIN);
//            style.setRightBorderColor(IndexedColors.GREY_40_PERCENT.getIndex());
//            style.setBorderLeft(CellStyle.BORDER_THIN);
//            style.setLeftBorderColor(IndexedColors.GREY_40_PERCENT.getIndex());
//            style.setBorderTop(CellStyle.BORDER_THIN);
//            style.setTopBorderColor(IndexedColors.GREY_40_PERCENT.getIndex());
//            style.setBorderBottom(CellStyle.BORDER_THIN);
//            style.setBottomBorderColor(IndexedColors.GREY_40_PERCENT.getIndex());
//            styles.put("redFont", style);
//
//            org.apache.poi.ss.usermodel.Font greenFont = wb.createFont();
//            greenFont.setFontHeightInPoints((short)11);
//            greenFont.setColor(IndexedColors.WHITE.getIndex());
//            style = wb.createCellStyle();
//            style.setAlignment(CellStyle.ALIGN_CENTER);
//            style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
//            style.setFillForegroundColor(IndexedColors.GREEN.getIndex());
//            style.setFillPattern(CellStyle.SOLID_FOREGROUND);
//            style.setFont(greenFont);
//            style.setWrapText(true);
//            style.setBorderRight(CellStyle.BORDER_THIN);
//            style.setRightBorderColor(IndexedColors.GREY_40_PERCENT.getIndex());
//            style.setBorderLeft(CellStyle.BORDER_THIN);
//            style.setLeftBorderColor(IndexedColors.GREY_40_PERCENT.getIndex());
//            style.setBorderTop(CellStyle.BORDER_THIN);
//            style.setTopBorderColor(IndexedColors.GREY_40_PERCENT.getIndex());
//            style.setBorderBottom(CellStyle.BORDER_THIN);
//            style.setBottomBorderColor(IndexedColors.GREY_40_PERCENT.getIndex());
//            styles.put("greenFont", style);
//
//            org.apache.poi.ss.usermodel.Font yellowFont = wb.createFont();
//            yellowFont.setFontHeightInPoints((short)11);
//            yellowFont.setColor(IndexedColors.WHITE.getIndex());
//            style = wb.createCellStyle();
//            style.setAlignment(CellStyle.ALIGN_CENTER);
//            style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
//            style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
//            style.setFillPattern(CellStyle.SOLID_FOREGROUND);
//            style.setFont(yellowFont);
//            style.setWrapText(true);
//            style.setBorderRight(CellStyle.BORDER_THIN);
//            style.setRightBorderColor(IndexedColors.GREY_40_PERCENT.getIndex());
//            style.setBorderLeft(CellStyle.BORDER_THIN);
//            style.setLeftBorderColor(IndexedColors.GREY_40_PERCENT.getIndex());
//            style.setBorderTop(CellStyle.BORDER_THIN);
//            style.setTopBorderColor(IndexedColors.GREY_40_PERCENT.getIndex());
//            style.setBorderBottom(CellStyle.BORDER_THIN);
//            style.setBottomBorderColor(IndexedColors.GREY_40_PERCENT.getIndex());
//            styles.put("yellowFont", style);
//
//            style = wb.createCellStyle();
//            style.setAlignment(CellStyle.ALIGN_LEFT);
//            style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
//            style.setWrapText(true);
//            style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
//            style.setFillPattern(CellStyle.SOLID_FOREGROUND);
//            style.setBorderRight(CellStyle.BORDER_THIN);
//            style.setRightBorderColor(IndexedColors.GREY_40_PERCENT.getIndex());
//            style.setBorderLeft(CellStyle.BORDER_THIN);
//            style.setLeftBorderColor(IndexedColors.GREY_40_PERCENT.getIndex());
//            style.setBorderTop(CellStyle.BORDER_THIN);
//            style.setTopBorderColor(IndexedColors.GREY_40_PERCENT.getIndex());
//            style.setBorderBottom(CellStyle.BORDER_THIN);
//            style.setBottomBorderColor(IndexedColors.GREY_40_PERCENT.getIndex());
//            styles.put("cellSub", style);
//
//            style = wb.createCellStyle();
//            style.setAlignment(CellStyle.ALIGN_CENTER);
//            style.setWrapText(true);
//            style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
//            style.setFillPattern(CellStyle.SOLID_FOREGROUND);
//            style.setBorderRight(CellStyle.BORDER_THIN);
//            style.setRightBorderColor(IndexedColors.GREY_40_PERCENT.getIndex());
//            style.setBorderLeft(CellStyle.BORDER_THIN);
//            style.setLeftBorderColor(IndexedColors.GREY_40_PERCENT.getIndex());
//            style.setBorderTop(CellStyle.BORDER_THIN);
//            style.setTopBorderColor(IndexedColors.GREY_40_PERCENT.getIndex());
//            style.setBorderBottom(CellStyle.BORDER_THIN);
//            style.setBottomBorderColor(IndexedColors.GREY_40_PERCENT.getIndex());
//            styles.put("cellSubCenter", style);
//
//            style = wb.createCellStyle();
//            style.setAlignment(CellStyle.ALIGN_LEFT);
//            style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
//            style.setWrapText(true);
//            style.setBorderRight(CellStyle.BORDER_THIN);
//            style.setRightBorderColor(IndexedColors.GREY_40_PERCENT.getIndex());
//            style.setBorderLeft(CellStyle.BORDER_THIN);
//            style.setLeftBorderColor(IndexedColors.GREY_40_PERCENT.getIndex());
//            style.setBorderTop(CellStyle.BORDER_THIN);
//            style.setTopBorderColor(IndexedColors.GREY_40_PERCENT.getIndex());
//            style.setBorderBottom(CellStyle.BORDER_THIN);
//            style.setBottomBorderColor(IndexedColors.GREY_40_PERCENT.getIndex());
//            styles.put("cell", style);
//
//            style = wb.createCellStyle();
//            style.setAlignment(CellStyle.ALIGN_CENTER);
//            style.setWrapText(true);
//            style.setBorderRight(CellStyle.BORDER_THIN);
//            style.setRightBorderColor(IndexedColors.GREY_40_PERCENT.getIndex());
//            style.setBorderLeft(CellStyle.BORDER_THIN);
//            style.setLeftBorderColor(IndexedColors.GREY_40_PERCENT.getIndex());
//            style.setBorderTop(CellStyle.BORDER_THIN);
//            style.setTopBorderColor(IndexedColors.GREY_40_PERCENT.getIndex());
//            style.setBorderBottom(CellStyle.BORDER_THIN);
//            style.setBottomBorderColor(IndexedColors.GREY_40_PERCENT.getIndex());
//            styles.put("cellCenter", style);
//
//            style = wb.createCellStyle();
//            style.setAlignment(CellStyle.ALIGN_RIGHT);
//            style.setWrapText(true);
//            style.setBorderRight(CellStyle.BORDER_THIN);
//            style.setRightBorderColor(IndexedColors.GREY_40_PERCENT.getIndex());
//            style.setBorderLeft(CellStyle.BORDER_THIN);
//            style.setLeftBorderColor(IndexedColors.GREY_40_PERCENT.getIndex());
//            style.setBorderTop(CellStyle.BORDER_THIN);
//            style.setTopBorderColor(IndexedColors.GREY_40_PERCENT.getIndex());
//            style.setBorderBottom(CellStyle.BORDER_THIN);
//            style.setBottomBorderColor(IndexedColors.GREY_40_PERCENT.getIndex());
//            style.setDataFormat(wb.createDataFormat().getFormat(BuiltinFormats.getBuiltinFormat(1) ));
//            styles.put("cellInteger", style);
//
//            style = wb.createCellStyle();
//            style.setAlignment(CellStyle.ALIGN_RIGHT);
//            style.setWrapText(true);
//            style.setBorderRight(CellStyle.BORDER_THIN);
//            style.setRightBorderColor(IndexedColors.GREY_40_PERCENT.getIndex());
//            style.setBorderLeft(CellStyle.BORDER_THIN);
//            style.setLeftBorderColor(IndexedColors.GREY_40_PERCENT.getIndex());
//            style.setBorderTop(CellStyle.BORDER_THIN);
//            style.setTopBorderColor(IndexedColors.GREY_40_PERCENT.getIndex());
//            style.setBorderBottom(CellStyle.BORDER_THIN);
//            style.setBottomBorderColor(IndexedColors.GREY_40_PERCENT.getIndex());
//            style.setDataFormat(wb.createDataFormat().getFormat(BuiltinFormats.getBuiltinFormat(4) ));
//            styles.put("cellDecimal", style);
//
//            style = wb.createCellStyle();
//            style.setAlignment(CellStyle.ALIGN_RIGHT);
//            style.setWrapText(true);
//            style.setBorderRight(CellStyle.BORDER_THIN);
//            style.setRightBorderColor(IndexedColors.GREY_40_PERCENT.getIndex());
//            style.setBorderLeft(CellStyle.BORDER_THIN);
//            style.setLeftBorderColor(IndexedColors.GREY_40_PERCENT.getIndex());
//            style.setBorderTop(CellStyle.BORDER_THIN);
//            style.setTopBorderColor(IndexedColors.GREY_40_PERCENT.getIndex());
//            style.setBorderBottom(CellStyle.BORDER_THIN);
//            style.setBottomBorderColor(IndexedColors.GREY_40_PERCENT.getIndex());
//            style.setDataFormat(wb.createDataFormat().getFormat("R$ ##,##0.00"));
//            styles.put("cellMoney", style);
//
//            style = wb.createCellStyle();
//            style.setAlignment(CellStyle.ALIGN_CENTER);
//            style.setWrapText(true);
//            style.setBorderRight(CellStyle.BORDER_THIN);
//            style.setRightBorderColor(IndexedColors.GREY_40_PERCENT.getIndex());
//            style.setBorderLeft(CellStyle.BORDER_THIN);
//            style.setLeftBorderColor(IndexedColors.GREY_40_PERCENT.getIndex());
//            style.setBorderTop(CellStyle.BORDER_THIN);
//            style.setTopBorderColor(IndexedColors.GREY_40_PERCENT.getIndex());
//            style.setBorderBottom(CellStyle.BORDER_THIN);
//            style.setBottomBorderColor(IndexedColors.GREY_40_PERCENT.getIndex());
//            style.setDataFormat(wb.createDataFormat().getFormat("dd-mm-yyyy"));
//            styles.put("cellDate", style);
//
//            style = wb.createCellStyle();
//            style.setAlignment(CellStyle.ALIGN_CENTER);
//            style.setWrapText(true);
//            style.setBorderRight(CellStyle.BORDER_THIN);
//            style.setRightBorderColor(IndexedColors.GREY_40_PERCENT.getIndex());
//            style.setBorderLeft(CellStyle.BORDER_THIN);
//            style.setLeftBorderColor(IndexedColors.GREY_40_PERCENT.getIndex());
//            style.setBorderTop(CellStyle.BORDER_THIN);
//            style.setTopBorderColor(IndexedColors.GREY_40_PERCENT.getIndex());
//            style.setBorderBottom(CellStyle.BORDER_THIN);
//            style.setBottomBorderColor(IndexedColors.GREY_40_PERCENT.getIndex());
//            style.setDataFormat(wb.createDataFormat().getFormat("0.00%"));
//            styles.put("cellPercent", style);
//
//            style = wb.createCellStyle();
//            monthFont.setFontHeightInPoints((short)10);
//            monthFont.setColor(IndexedColors.WHITE.getIndex());
//            style.setAlignment(CellStyle.ALIGN_CENTER);
//            style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
//            style.setFillForegroundColor(IndexedColors.GREEN.getIndex());
//            style.setFillPattern(CellStyle.SOLID_FOREGROUND);
//            style.setDataFormat(wb.createDataFormat().getFormat("R$ ##,##0.00"));
//            style.setFont(monthFont);
//            styles.put("formula", style);
//
//            style = wb.createCellStyle();
//            style.setAlignment(CellStyle.ALIGN_CENTER);
//            style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
//            style.setFillForegroundColor(IndexedColors.GREY_40_PERCENT.getIndex());
//            style.setFillPattern(CellStyle.SOLID_FOREGROUND);
//            style.setDataFormat(wb.createDataFormat().getFormat("0.00"));
//            styles.put("formula_2", style);
//
//            return styles;
//        }
//
//
//        public static Font getFontWhite(int tamanho, boolean bold) {
//            if (bold) {
//                return FontFactory.getFont(FontFactory.HELVETICA, tamanho, Font.BOLD, new Color(255, 255, 255));
//            } else {
//                return FontFactory.getFont(FontFactory.HELVETICA, tamanho, Font.NORMAL, new Color(255, 255, 255));
//            }
//        }
//
//        public static Font getFont(int tamanho, boolean bold, boolean italic) {
//            if (bold) {
//                return FontFactory.getFont(FontFactory.HELVETICA, tamanho, Font.BOLDITALIC, new Color(0, 0, 0));
//            } else if (italic) {
//                return FontFactory.getFont(FontFactory.HELVETICA, tamanho, Font.ITALIC, new Color(0, 0, 0));
//            } else {
//                return FontFactory.getFont(FontFactory.HELVETICA, tamanho, Font.NORMAL, new Color(0, 0, 0));
//            }
//        }
//
//        public static PdfPCell getCellTitle(Paragraph paragraph, int colspan) {
//            PdfPCell cell = new PdfPCell(paragraph);
//            cell.setHorizontalAlignment(Element.ALIGN_LEFT);
//            cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
//            cell.setBorderColor(Color.lightGray);
//            cell.setBorder(PdfPCell.NO_BORDER);
//            cell.setBackgroundColor(Color.getHSBColor(0, 0, Float.valueOf("0.981")));
//            cell.setPadding(4f);
//            cell.setColspan(colspan);
//
//            return cell;
//        }
//
//        public static PdfPCell getCellTitleGray(Paragraph paragraph, int colspan) {
//            PdfPCell cell = new PdfPCell(paragraph);
//            cell.setHorizontalAlignment(Element.ALIGN_LEFT);
//            cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
//            cell.setBorder(PdfPCell.NO_BORDER);
//            cell.setBackgroundColor(Color.getHSBColor(0, 0, Float.valueOf("0.961")));
//            cell.setPadding(4f);
//            cell.setColspan(colspan);
//
//            return cell;
//        }
//
//        public static PdfPCell getCellTitleBlackCenterBorder(Paragraph paragraph, int colspan) {
//            PdfPCell cell = new PdfPCell(paragraph);
//            cell.setHorizontalAlignment(Element.ALIGN_CENTER);
//            cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
//            cell.setBorderColor(Color.black);
//            cell.setBorderWidth(0.7f);
//            cell.setBackgroundColor(Color.black);
//            cell.setPadding(4f);
//            cell.setColspan(colspan);
//
//            return cell;
//        }
//
//        public static PdfPCell getCellTitleGrayCenterBorder(Paragraph paragraph, int colspan) {
//            PdfPCell cell = new PdfPCell(paragraph);
//            cell.setHorizontalAlignment(Element.ALIGN_CENTER);
//            cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
//            cell.setBorderColor(Color.lightGray);
//            cell.setBorderWidth(0.7f);
//            cell.setBackgroundColor(Color.getHSBColor(0, 0, Float.valueOf("0.961")));
//            cell.setPadding(4f);
//            cell.setColspan(colspan);
//
//            return cell;
//        }
//
//        public static PdfPCell getCellTitleGrayDireitaBorder(Paragraph paragraph, int colspan) {
//            PdfPCell cell = new PdfPCell(paragraph);
//            cell.setHorizontalAlignment(Element.ALIGN_RIGHT);
//            cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
//            cell.setBorderColor(Color.lightGray);
//            cell.setBorderWidth(0.7f);
//            cell.setBackgroundColor(Color.getHSBColor(0, 0, Float.valueOf("0.961")));
//            cell.setPadding(4f);
//            cell.setColspan(colspan);
//
//            return cell;
//        }
//
//        public static PdfPCell getCellLabel(Paragraph paragraph, int colspan) {
//            PdfPCell cell = new PdfPCell(paragraph);
//            cell.setHorizontalAlignment(Element.ALIGN_LEFT);
//            cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
//            cell.setBorder(PdfPCell.NO_BORDER);
//            cell.setColspan(colspan);
//
//            return cell;
//        }
//
//        public static PdfPCell getCellLabelBorder(Paragraph paragraph, int colspan) {
//            PdfPCell cell = new PdfPCell(paragraph);
//            cell.setHorizontalAlignment(Element.ALIGN_LEFT);
//            cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
//            cell.setBorderColor(Color.lightGray);
//            cell.setBorderWidth(0.5f);
//            cell.setPadding(4f);
//            cell.setColspan(colspan);
//
//            return cell;
//        }
//
//        public static PdfPCell getCellLabelBorderLeft(Paragraph paragraph, int colspan, Color color) {
//            PdfPCell cell = new PdfPCell(paragraph);
//            cell.setHorizontalAlignment(Element.ALIGN_LEFT);
//            cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
//            cell.setBorderColor(Color.lightGray);
//            cell.setBorderWidth(0.5f);
//            cell.setPadding(4f);
//            cell.setColspan(colspan);
//
//            cell.setBorderWidthLeft(0f);
//            cell.setBorderColorLeft(color);
//            return cell;
//        }
//
//        public static PdfPCell getCellLabelCenter(Paragraph paragraph, int colspan) {
//            PdfPCell cell = new PdfPCell(paragraph);
//            cell.setHorizontalAlignment(Element.ALIGN_LEFT);
//            cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
//            cell.setBorder(PdfPCell.NO_BORDER);
//            cell.setColspan(colspan);
//
//            return cell;
//        }
//
//        public static PdfPCell getCellLabelCenterBorder(Paragraph paragraph, int colspan) {
//            PdfPCell cell = new PdfPCell(paragraph);
//            cell.setHorizontalAlignment(Element.ALIGN_CENTER);
//            cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
//            cell.setBorderColor(Color.lightGray);
//            cell.setBorderWidth(0.5f);
//            cell.setPadding(4f);
//            cell.setColspan(colspan);
//
//            return cell;
//        }
//
//        public static PdfPCell getCellLabelRightBorder(Paragraph paragraph, int colspan) {
//            PdfPCell cell = new PdfPCell(paragraph);
//            cell.setHorizontalAlignment(Element.ALIGN_RIGHT);
//            cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
//            cell.setBorderColor(Color.lightGray);
//            cell.setBorderWidth(0.5f);
//            cell.setPadding(4f);
//            cell.setColspan(colspan);
//
//            return cell;
//        }
//
//        public static PdfPCell getCellLabelNoBorder(Paragraph paragraph, int colspan) {
//            PdfPCell cell = new PdfPCell(paragraph);
//
//            cell.setHorizontalAlignment(Element.ALIGN_LEFT);
//            cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
//            cell.setBorder(PdfPCell.NO_BORDER);
//            cell.setColspan(colspan);
//            cell.setPadding(5f);
//
//            return cell;
//        }
//
//        public static PdfPCell getCellLabelCenterBorderWhite(Paragraph paragraph, int colspan) {
//            PdfPCell cell = new PdfPCell(paragraph);
//            cell.setHorizontalAlignment(Element.ALIGN_LEFT);
//            cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
//
//            cell.setBorderWidthTop(0.5f);
//            cell.setBorderColorTop(Color.white);
//            cell.setBorderWidthBottom(0.5f);
//            cell.setBorderColorBottom(Color.white);
//            cell.setBorderWidthLeft(0.5f);
//            cell.setBorderColorLeft(Color.lightGray);
//            cell.setBorderWidthRight(0.5f);
//            cell.setBorderColorRight(Color.lightGray);
//            cell.setPadding(4f);
//
//            cell.setColspan(colspan);
//
//            return cell;
//        }
//
//        public static PdfPCell getCellLabelCenterBorderWhite(Image image, int colspan) {
//            PdfPCell cell = new PdfPCell(image);
//            cell.setHorizontalAlignment(Element.ALIGN_LEFT);
//            cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
//
//            cell.setBorderWidthTop(0.5f);
//            cell.setBorderColorTop(Color.white);
//            cell.setBorderWidthBottom(0.5f);
//            cell.setBorderColorBottom(Color.white);
//            cell.setBorderWidthLeft(0.5f);
//            cell.setBorderColorLeft(Color.lightGray);
//            cell.setBorderWidthRight(0.5f);
//            cell.setBorderColorRight(Color.lightGray);
//            cell.setPadding(4f);
//
//            cell.setColspan(colspan);
//
//            return cell;
//        }
//
//        public static PdfPCell getCellLabelNoBorder(Image image, int colspan) {
//            PdfPCell cell = new PdfPCell(image);
//
//            cell.setHorizontalAlignment(Element.ALIGN_LEFT);
//            cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
//            cell.setBorder(PdfPCell.NO_BORDER);
//            cell.setPadding(5f);
//            cell.setColspan(colspan);
//
//            return cell;
//        }
//
//        public static PdfPCell getCellEspaco(int colspan) {
//            PdfPCell cell = new PdfPCell(new Paragraph(""));
//            cell.setBorder(PdfPCell.NO_BORDER);
//            cell.setColspan(colspan);
//
//            return cell;
//        }
//
//
//    }
}