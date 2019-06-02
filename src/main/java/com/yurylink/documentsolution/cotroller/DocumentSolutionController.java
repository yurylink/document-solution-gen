package com.yurylink.documentsolution.cotroller;

import com.yurylink.documentsolution.converter.XlsxToPdfConverter;
import com.yurylink.documentsolution.creator.BasicExcelArchiveDto;
import com.yurylink.documentsolution.creator.impl.ExcelBasicCreator;
import org.springframework.core.io.InputStreamResource;
import org.springframework.core.io.Resource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

@RestController
@RequestMapping("/document")
public class DocumentSolutionController {

    @GetMapping("/xlsx")
    public ResponseEntity<Resource> getExampleDocAsXlsx(
            /*final @RequestParam(value = "date", required = false) @DateTimeFormat(pattern = "yyyy-MM-dd") Date date,
            final @RequestParam(value = "ID", required = false) String ids,
            final @RequestParam(value = "reportName", required = false) String name*/){

        List<String> values = new ArrayList<String>();
        for (int i = 0; i < 5 ; i++) {
            values.add("Values " + i);
        }
        BasicExcelArchiveDto dto = new BasicExcelArchiveDto();
        dto.setName("Example Archive");
        dto.setValues(values);

        ExcelBasicCreator create = new ExcelBasicCreator(dto);

        try {
            ByteArrayInputStream in = create.build();
            final String fileName = dto.getName() + ".xlsx";

            HttpHeaders header = new HttpHeaders();
            header.add("Content-Disposition", "attachment; filename=" + fileName);

            return ResponseEntity.ok().headers(header).body(new InputStreamResource(in));
        } catch (IOException e) {
            e.printStackTrace();
            return ResponseEntity.badRequest().build();
        }
    }

    @GetMapping("/pdf")
    public ResponseEntity<Resource> getExampleDocAsPdf(
            /*final @RequestParam(value = "date", required = false) @DateTimeFormat(pattern = "yyyy-MM-dd") Date date,
            final @RequestParam(value = "ID", required = false) String ids,
            final @RequestParam(value = "reportName", required = false) String name*/){

        List<String> values = new ArrayList<String>();
        for (int i = 0; i < 5 ; i++) {
            values.add("Values " + i);
        }
        BasicExcelArchiveDto dto = new BasicExcelArchiveDto();
        dto.setName("Example Archive");
        dto.setValues(values);

        ExcelBasicCreator create = new ExcelBasicCreator(dto);

        try {
            ByteArrayInputStream xlsxIn = create.build();
            final String fileName = dto.getName() + ".pdf" +
                    "";

            ByteArrayInputStream pdfIn =
                    new XlsxToPdfConverter(xlsxIn, create.getColumnsQuantity()).
                            convertXlsxToPdf();

            HttpHeaders header = new HttpHeaders();
            header.add("Content-Disposition", "attachment; filename=" + fileName);

            return ResponseEntity.ok().headers(header).body(new InputStreamResource(pdfIn));
        } catch (Exception e) {
            e.printStackTrace();
            return ResponseEntity.badRequest().build();
        }
    }

}