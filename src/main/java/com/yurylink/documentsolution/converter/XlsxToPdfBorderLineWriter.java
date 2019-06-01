package com.yurylink.documentsolution.converter;

import com.itextpdf.text.Rectangle;
import com.itextpdf.text.pdf.PdfContentByte;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPCellEvent;
import com.itextpdf.text.pdf.PdfPTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

public class XlsxToPdfBorderLineWriter implements PdfPCellEvent {

    private final XSSFCellStyle cellStyle;

    public XlsxToPdfBorderLineWriter(XSSFCellStyle cellStyle) {
        this.cellStyle = cellStyle;
    }

    public void cellLayout(PdfPCell cell, Rectangle position,
                           PdfContentByte[] canvases) {
        PdfContentByte canvas = canvases[PdfPTable.LINECANVAS];
        canvas.saveState();
//        canvas.setLineDash(0, 4, 2);
//        canvas.setLineDash(0);
        if (hasTopBorder()) {
            canvas.moveTo(position.getRight(), position.getTop());
            canvas.lineTo(position.getLeft(), position.getTop());
        }
        if (hasBottomBorder()) {
            canvas.moveTo(position.getRight(), position.getBottom());
            canvas.lineTo(position.getLeft(), position.getBottom());
        }
        if (hasRightBorder()) {
            canvas.moveTo(position.getRight(), position.getTop());
            canvas.lineTo(position.getRight(), position.getBottom());
        }
        if (hasLeftBorder()) {
            canvas.moveTo(position.getLeft(), position.getTop());
            canvas.lineTo(position.getLeft(), position.getBottom());
        }
        canvas.stroke();
        canvas.restoreState();
    }

    private boolean hasTopBorder(){
        if(cellStyle.getBorderTop() != XSSFCellStyle.BORDER_NONE)
            return true;
        return false;
    }

    private boolean hasBottomBorder(){
        if(cellStyle.getBorderBottom() != XSSFCellStyle.BORDER_NONE)
            return true;
        return false;
    }

    private boolean hasLeftBorder(){
        if(cellStyle.getBorderLeft() != XSSFCellStyle.BORDER_NONE)
            return true;
        return false;
    }

    private boolean hasRightBorder(){
        if(cellStyle.getBorderRight() != XSSFCellStyle.BORDER_NONE)
            return true;
        return false;
    }

}
