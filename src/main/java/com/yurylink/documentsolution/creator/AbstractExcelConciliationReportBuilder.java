
package com.yurylink.documentsolution.creator;

import com.yurylink.documentsolution.builder.ExcelReportBuilder;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.util.List;

/**
 * Implementation of ExcellReportBuilder, expect that subclasses implements the 3 abstract methos and the constructor
 *
 * @author Wesley Yuri
 * @param <T>
 */
public abstract class AbstractExcelConciliationReportBuilder<T> {

	/**
	 * The Excell Title
	 * 
	 * @return
	 */
	public abstract String getTitleText();

	/**
	 * The Subtitle with the date
	 * 
	 * @return
	 */
	public abstract String getSubTitleText();

	/**
	 * A list of string for the headers
	 * 
	 * @return
	 */
	public abstract List<String> getHeader();

	protected ExcelReportBuilder builder;
	protected Integer columnsQuantity;
	protected T dto;

	public AbstractExcelConciliationReportBuilder(final T dto) {
		this.dto = dto;
		this.columnsQuantity = getHeader().size();
		this.builder = ExcelReportBuilder.newInstance(columnsQuantity);
	}

	protected abstract void buildBody();

	protected void buildHeader() {
		builder.addTitle(getTitleText()).addSubTitle(getSubTitleText()).addHeaderColumns(getHeader());
	}

	public ByteArrayInputStream build() throws IOException {
		buildHeader();
		buildBody();
		return builder.build();
	}

	public Integer getColumnsQuantity() {
		return columnsQuantity;
	}

	public static String getTranslatedStatus(final String originalStatus) {
		if (originalStatus.equalsIgnoreCase("NOT_RECONCILED")) {
			return "Não conciliado";
		}

		if (originalStatus.equalsIgnoreCase("RECONCILED")) {
			return "Conciliado";
		}
		return originalStatus;
	}
}
