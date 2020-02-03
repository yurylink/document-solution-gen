package com.yurylink.file;

import org.beanio.annotation.Field;
import org.beanio.annotation.Record;
import org.beanio.builder.Align;

import java.util.Date;

@Record
public class FileLineRegisterBean {

    @Field(ordinal = 0, length = 14, align = Align.RIGHT, format = "00000000000000", padding = 0, trim = true, rid = true)
    private Long cpfcnpj;

    @Field(ordinal = 1, length = 6, align = Align.RIGHT, format = "000000", padding = 0, trim = true)
    private Long idConta;

    @Field(ordinal = 2, length = 10)
    private String estabelecimento;

    @Field(ordinal = 3, length = 8, format = "yyyyMMdd")
    private Date dataTransacao;

    @Field(ordinal = 4, length = 8, format = "yyyyMMdd")
    private Date dataLiquidacao;

    @Field(ordinal = 5, length = 10, padding = 0, align = Align.RIGHT, format = "0000000000")
    private Long valor;

    @Field(ordinal = 6, length = 8, type = Date.class, handlerName = "dateTimeZero", handlerClass = CustomDateTimeHandler.class)
    private Date dataProcessamento;

    @Field(ordinal = 7, length = 3,  align = Align.RIGHT, format = "000", padding = 0, trim = true, defaultValue = "000")
    private Long statusProcessamento;

    @Field(ordinal = 8, length = 40, align = Align.RIGHT, format = "0000000000000000000000000000000000000000", padding = 0, trim = true)
    private Long transferId;

    public FileLineRegisterBean() {
    }

    public FileLineRegisterBean(Long cpfcnpj, Long idConta, String estabelecimento, Date dataTransacao, Date dataLiquidacao, Long valor, Date dataProcessamento, Long statusProcessamento, Long transferId) {
        this.cpfcnpj = cpfcnpj;
        this.idConta = idConta;
        this.estabelecimento = estabelecimento;
        this.dataTransacao = dataTransacao;
        this.dataLiquidacao = dataLiquidacao;
        this.valor = valor;
        this.dataProcessamento = dataProcessamento;
        this.statusProcessamento = statusProcessamento;
        this.transferId = transferId;
    }

    public Long getCpfcnpj() {
        return cpfcnpj;
    }

    public void setCpfcnpj(Long cpfcnpj) {
        this.cpfcnpj = cpfcnpj;
    }

    public Long getIdConta() {
        return idConta;
    }

    public void setIdConta(Long idConta) {
        this.idConta = idConta;
    }

    public String getEstabelecimento() {
        return estabelecimento;
    }

    public void setEstabelecimento(String estabelecimento) {
        this.estabelecimento = estabelecimento;
    }

    public Date getDataTransacao() {
        return dataTransacao;
    }

    public void setDataTransacao(Date dataTransacao) {
        this.dataTransacao = dataTransacao;
    }

    public Date getDataLiquidacao() {
        return dataLiquidacao;
    }

    public void setDataLiquidacao(Date dataLiquidacao) {
        this.dataLiquidacao = dataLiquidacao;
    }

    public Long getValor() {
        return valor;
    }

    public void setValor(Long valor) {
        this.valor = valor;
    }

    public Date getDataProcessamento() {
        return dataProcessamento;
    }

    public void setDataProcessamento(Date dataProcessamento) {
        this.dataProcessamento = dataProcessamento;
    }

    public Long getStatusProcessamento() {
        return statusProcessamento;
    }

    public void setStatusProcessamento(Long statusProcessamento) {
        this.statusProcessamento = statusProcessamento;
    }

    public Long getTransferId() {
        return transferId;
    }

    public void setTransferId(Long transferId) {
        this.transferId = transferId;
    }


}