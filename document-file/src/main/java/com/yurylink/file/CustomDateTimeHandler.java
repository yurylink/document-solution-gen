package com.yurylink.file;

import jdk.nashorn.internal.runtime.ParserException;
import org.beanio.types.TypeConversionException;
import org.beanio.types.TypeHandler;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

public class CustomDateTimeHandler implements TypeHandler {
    private SimpleDateFormat dateFormat = new SimpleDateFormat("yyyyMMdd");

    private static final Logger LOGGER = LoggerFactory.getLogger(CustomDateTimeHandler.class);

    @Override
    public Object parse(String text) throws TypeConversionException {
        try {
            if(text == null || text.isEmpty())
                throw new ParserException("Data nulla");
            return dateFormat.parse(text);
        } catch (ParseException e) {
            LOGGER.warn("Erro ao passar a data " + text );
            return null;
        }
    }

    @Override
    public String format(Object value) {
        try {
            if(value == null)
                throw new Exception("Data nulla");
            return dateFormat.format(value);
        } catch (Exception e) {
            LOGGER.info("Campo data nullo, será utizado valor padrão.");
            return "00000000";
        }
    }

    @Override
    public Class<?> getType() {
        return Date.class;
    }
}