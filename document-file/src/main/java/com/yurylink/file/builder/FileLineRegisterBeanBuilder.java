package com.yurylink.file.builder;

import com.yurylink.file.FileLineRegisterBean;

import java.util.Date;

public final class FileLineRegisterBeanBuilder {
    private Long cpfcnpj;
    private Long idConta;
    private String estabelecimento;
    private Date dataTransacao;
    private Date dataLiquidacao;
    private Long valor;
    private Date dataProcessamento;
    private Long statusProcessamento;
    private Long transferId;

    private FileLineRegisterBeanBuilder() {
    }

    public static FileLineRegisterBeanBuilder create() {
        return new FileLineRegisterBeanBuilder();
    }

    public FileLineRegisterBeanBuilder setCpfcnpj(Long cpfcnpj) {
        this.cpfcnpj = cpfcnpj;
        return this;
    }

    public FileLineRegisterBeanBuilder setIdConta(Long idConta) {
        this.idConta = idConta;
        return this;
    }

    public FileLineRegisterBeanBuilder setEstabelecimento(String estabelecimento) {
        this.estabelecimento = estabelecimento;
        return this;
    }

    public FileLineRegisterBeanBuilder setDataTransacao(Date dataTransacao) {
        this.dataTransacao = dataTransacao;
        return this;
    }

    public FileLineRegisterBeanBuilder setDataLiquidacao(Date dataLiquidacao) {
        this.dataLiquidacao = dataLiquidacao;
        return this;
    }

    public FileLineRegisterBeanBuilder setValor(Long valor) {
        this.valor = valor;
        return this;
    }

    public FileLineRegisterBeanBuilder setDataProcessamento(Date dataProcessamento) {
        this.dataProcessamento = dataProcessamento;
        return this;
    }

    public FileLineRegisterBeanBuilder setStatusProcessamento(Long statusProcessamento) {
        this.statusProcessamento = statusProcessamento;
        return this;
    }

    public FileLineRegisterBeanBuilder setTransferId(Long transferId) {
        this.transferId = transferId;
        return this;
    }

    public FileLineRegisterBean build() {
        return new FileLineRegisterBean(cpfcnpj, idConta, estabelecimento, dataTransacao, dataLiquidacao, valor, dataProcessamento, statusProcessamento, transferId);
    }
}
