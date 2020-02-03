package com.yurylink.documentsolution.cotroller;

import com.yurylink.documentsolution.business.BeanioBussiness;
import org.apache.poi.util.IOUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.MediaType;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.bind.annotation.RestController;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;

@RestController
@RequestMapping("/beanio")
public class DocumentRWController {

    @Autowired
    private BeanioBussiness beanioBussiness;

    @GetMapping(
            value = "/get-file",
            produces = MediaType.APPLICATION_OCTET_STREAM_VALUE
    )
    public @ResponseBody byte[] getFile() throws IOException {
        File file = beanioBussiness.generateExampleFile();
        InputStream in = getClass()
                .getResourceAsStream(file.getAbsolutePath());
        return IOUtils.toByteArray(in);
    }


}
