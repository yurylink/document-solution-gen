package com.yurylink.documentsolution.business;

import com.yurylink.file.CustomDateTimeHandler;
import com.yurylink.file.FileLineRegisterBean;
import com.yurylink.file.builder.FileLineRegisterBeanBuilder;
import org.beanio.BeanReader;
import org.beanio.BeanWriter;
import org.beanio.StreamFactory;
import org.beanio.builder.FixedLengthParserBuilder;
import org.beanio.builder.StreamBuilder;
import org.springframework.stereotype.Component;

import java.io.File;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

@Component
public class BeanioBussiness {

    private BeanReader beanReader;
    private BeanWriter beanWriter;

    public File generateExampleFile(){
        StreamFactory factory = StreamFactory.newInstance();
        StreamBuilder builderCSV = new StreamBuilder("expBuilder")
                .format("fixedlength")
                .parser(new FixedLengthParserBuilder())
                .addTypeHandler("dateTimeZero", new CustomDateTimeHandler())
                .addRecord(FileLineRegisterBean.class);
        factory.define(builderCSV);

        File physicalFile = new File("src/main/resources/generateFile"+(new Date()).getTime()+".txt");
        beanWriter = factory.createWriter("conductor", physicalFile);
        List<FileLineRegisterBean> list = getMockList();

        getMockList().forEach(f -> {
            beanWriter.write(f);
        });
        beanWriter.close();
        return physicalFile;
    }

    private List<FileLineRegisterBean> getMockList(){
        List<FileLineRegisterBean> beans = new ArrayList<>();
        for (Integer i = 1; i < 11 ; i++) {
            beans.add(
                    FileLineRegisterBeanBuilder
                            .create()
                            .setCpfcnpj(i.longValue())
                            .setIdConta(i.longValue())
                            .setEstabelecimento("text"+i)
                            .setDataLiquidacao(new Date())
                            .setDataLiquidacao(new Date())
                            .setValor(i.longValue()*10)
                            .setDataProcessamento(new Date())
                            .setStatusProcessamento(i.longValue())
                            .setTransferId(i.longValue())
                            .build()
            );
        }
        return beans;
    }

    public File readeExampleFile(){
        StreamFactory factory = StreamFactory.newInstance();
        StreamBuilder streamBuilder = new StreamBuilder("conductor")
                .format("fixedlength")
                .parser(new FixedLengthParserBuilder())
                .addTypeHandler("dateTimeZero", new CustomDateTimeHandler())
                .addRecord(FileLineRegisterBean.class);
        factory.define(streamBuilder);

        File physicalFile = new File("src/main/resources/generateFile.txt");

        beanReader = factory.createReader("conductor", physicalFile);

        List<FileLineRegisterBean> list = getMockList();
        return physicalFile;
    }
}
