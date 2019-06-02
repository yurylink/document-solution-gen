
package com.yurylink.documentsolution.Utils;

import com.itextpdf.text.BaseColor;
import com.itextpdf.text.Font;
import com.itextpdf.text.FontFactory;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.springframework.stereotype.Component;

import java.awt.*;

public class ExcellReportConstants {

    public static final XSSFColor GREEN_ADIQ = new XSSFColor(new Color(0,222,178));
    public static final XSSFColor BLUE_ADIQ = new XSSFColor(new Color(184,246,233));
    public static final XSSFColor SOFT_GREY = new XSSFColor(new Color(206,206,210));
    public static final XSSFColor DARK_GREY = new XSSFColor(new Color(76,81,95));
    public static final XSSFColor WHITE_COLOR = new XSSFColor(Color.WHITE);
    public static final XSSFColor TITLE_BASE_COLOR = new XSSFColor(Color.LIGHT_GRAY); // new
    public static final XSSFColor LINE_SEPARATOR_COLOR = new XSSFColor(Color.GRAY);// new java.awt.Color(76, 81, 95));

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
}

