
package com.yurylink.documentsolution.converter;

import com.itextpdf.text. Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Element;
import com.itextpdf.text.Phrase;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;
import com.yurylink.documentsolution.builder.ExcelUtil;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.text.NumberFormat;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Locale;

/**
 * Receive a ByteArrayInputStream of the xlsx file and returns a ByteArrayInputStream of the converter documento
 *
 * @author Wesley Yuri
 */
public class XlsxToPdfConverter {

	private static final Logger LOGGER = LoggerFactory.getLogger(XlsxToPdfConverter.class);
	private static final NumberFormat NF = NumberFormat.getCurrencyInstance(new Locale("da", "DK"));
	private List<CellRangeAddress> listOfMergedCells;
	private final Document documentPdf;
	private final XSSFWorkbook workbook;
	private final Integer columnsQuantity;
	private String titleText;

	/**
	 * Receive a ByteArrayInputStream of the xlsx file and returns a ByteArrayInputStream of the converter documento
	 * 
	 * @param xlsxIS
	 * @param columnsQuantity
	 * @return
	 */
	public XlsxToPdfConverter(final ByteArrayInputStream xlsxIS, final Integer columnsQuantity) throws Exception {
		documentPdf = new Document();
		workbook = new XSSFWorkbook(xlsxIS);
		this.columnsQuantity = columnsQuantity;
		listOfMergedCells = null;
		titleText = "";
	}

	public ByteArrayInputStream convertXlsxToPdf() throws Exception {
		try {
			final ByteArrayOutputStream outputPdfStream = new ByteArrayOutputStream();
			PdfWriter.getInstance(documentPdf, outputPdfStream);
			documentPdf.open();

			final Iterator<XSSFSheet> sheetIterator = workbook.iterator();
			final PdfPTable pdfTable = new PdfPTable(columnsQuantity);

			// SHEET
			while (sheetIterator.hasNext()) {
				documentPdf.newPage();
				final XSSFSheet currentSheet = sheetIterator.next();
				listOfMergedCells = getAllMergedRegionFromSheet(currentSheet);

				// ROW
				for (int i = 0; i <= currentSheet.getLastRowNum(); i++) {
					final XSSFRow currentRow = currentSheet.getRow(i);
					pdfTable.completeRow();

					// EMPTY ROLL && ROW FULL OF EMPTY CELLS
					if (currentRow.getLastCellNum() < 0 || currentRow.getLastCellNum() > 0 && isEmptyRow(currentRow)) {
						final PdfPCell nCell = new PdfPCell(new Phrase());
						nCell.setColspan(columnsQuantity);
						continue;
					}

					// CELL
					for (int j = 0; j < columnsQuantity; j++) {
						final XSSFCell currentCell = currentRow.getCell(j);
						final PdfPCell cell = getPdfCellValueFromXlsxCell(currentCell);
						if (cell == null) {
							continue;
						}
						if ((i == 0 || i == 1) && j == 0) {
							titleText += cell.getPhrase().getContent();
						}
						handleMergedAreas(currentCell, cell);
						pdfTable.addCell(cell);
					}
				}
			}
			documentPdf.addTitle(titleText);
			documentPdf.add(pdfTable);
			documentPdf.close();
			return new ByteArrayInputStream(outputPdfStream.toByteArray());
		} catch (final DocumentException e) {
			LOGGER.error("Error on handle PDF document: " + e.getMessage());
			e.printStackTrace();
			throw new Exception(e);
		} catch (final Exception e) {
			LOGGER.error("Unespected error during the file conversion: " + e.getMessage());
			e.printStackTrace();
			throw e;
		}
	}

	private Boolean isEmptyRow(final XSSFRow currentRow) {
		for (int i = 0; i < currentRow.getLastCellNum(); i++) {
			final XSSFCell currentCell = currentRow.getCell(i);
			if (XSSFCell.CELL_TYPE_STRING == currentCell.getCellType()) {
				if (!currentCell.getStringCellValue().isEmpty()) {
					return false;
				}
			} else if (XSSFCell.CELL_TYPE_NUMERIC == currentCell.getCellType()) {
				if (!(Double.compare(0, currentCell.getNumericCellValue()) != 0)) {
					return false;
				}
			}
		}
		return true;
	}

	private List<CellRangeAddress> getAllMergedRegionFromSheet(final XSSFSheet sheet) {
		final List<CellRangeAddress> result = new ArrayList<>();
		for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
			result.add(sheet.getMergedRegion(i));
		}
		return result;
	}

	private PdfPCell getPdfCellValueFromXlsxCell(final XSSFCell currentCell) {
		PdfPCell cell = null;
		if (currentCell != null) {
			if (XSSFCell.CELL_TYPE_STRING == currentCell.getCellType()) {
				final Phrase nPhrase = new Phrase(currentCell.getStringCellValue(),
						ExcelUtil.getFont(8, currentCell.getCellStyle().getFont().getBold()));

				cell = new PdfPCell(nPhrase);
				cell.setHorizontalAlignment(PdfPCell.ALIGN_CENTER);
				cell.setVerticalAlignment(PdfPCell.ALIGN_CENTER);

			} else if (XSSFCell.CELL_TYPE_NUMERIC == currentCell.getCellType()) {
				final Phrase nPhrase = new Phrase(formatCurrency(currentCell.getNumericCellValue()),
						ExcelUtil.getFont(8, currentCell.getCellStyle().getFont().getBold()));

				cell = new PdfPCell(nPhrase);
				cell.setHorizontalAlignment(PdfPCell.ALIGN_CENTER);
				cell.setVerticalAlignment(PdfPCell.ALIGN_CENTER);

			}
		}
		return cell;
	}

	private void handleMergedAreas(final XSSFCell xlsxCell, final PdfPCell pdfPCell) {
		final CellRangeAddress range = getCellRangeAddress(xlsxCell);
		if (range != null) {
			pdfPCell.setColspan(range.getLastColumn() - range.getFirstColumn() + 1);
			pdfPCell.setRowspan(range.getLastRow() - range.getFirstRow() + 1);
			pdfPCell.setHorizontalAlignment(Element.ALIGN_CENTER);
		}
	}

	private CellRangeAddress getCellRangeAddress(final XSSFCell candidate) {
		return listOfMergedCells.stream()
				.filter(cellRangeAddress -> compareCellRangeAddressWithCell(candidate, cellRangeAddress))
				.findFirst()
				.orElse(null);
	}

	private Boolean compareCellRangeAddressWithCell(final XSSFCell cell, final CellRangeAddress cra) {
		return cra.getFirstColumn() == cell.getColumnIndex() && cra.getFirstRow() == cell.getRowIndex();
	}

	private static String formatCurrency(final Double unformatedValue) {
		return NF.format(unformatedValue).replace("kr ", "");
	}
}
