package com.yurylink.documentsolution.creator.impl;

import com.yurylink.documentsolution.creator.AbstractExcelConciliationReportBuilder;
import com.yurylink.documentsolution.creator.BasicExcelArchiveDto;

import java.util.ArrayList;
import java.util.List;

public class ExcelBasicCreator extends AbstractExcelConciliationReportBuilder<BasicExcelArchiveDto> {

    public ExcelBasicCreator(BasicExcelArchiveDto dto) {
        super(dto);
    }

    @Override
    public String getTitleText() {
        return dto.getName();
    }

    @Override
    public String getSubTitleText() {
        return "SubTitle example";
    }

    @Override
    public List<String> getHeader() {
        List<String> values = new ArrayList<String>();
        for (int i = 0; i < 5 ; i++) {
            values.add("Values " + i);
        }
        return values;
    }

    @Override
    protected void buildBody() {
        builder.newRow().
                addLineSeparatorSection("LineSeparator").
                newRow();
        for (String current :getHeader()) {
            builder.addCell(current);
        }
    }
}