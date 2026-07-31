package com.example.demo.service.impl;

import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.springframework.stereotype.Service;

import com.example.demo.service.IExcelToPdfService;
import com.lowagie.text.Document;
import com.lowagie.text.DocumentException;
import com.lowagie.text.Element;
import com.lowagie.text.Font;
import com.lowagie.text.FontFactory;
import com.lowagie.text.PageSize;
import com.lowagie.text.Phrase;
import com.lowagie.text.Rectangle;
import com.lowagie.text.pdf.PdfPCell;
import com.lowagie.text.pdf.PdfPTable;
import com.lowagie.text.pdf.PdfWriter;

import lombok.RequiredArgsConstructor;

@Service("exceltopdfserv")
public class ExcelToPdfServImpl implements IExcelToPdfService {
	
//	@Override
//    public byte[] convertExcelToPdf(InputStream excelStream) throws Exception {

//        try (Workbook workbook = new XSSFWorkbook(excelStream)) {
//
//            Sheet sheet = workbook.getSheetAt(0);
//
//            ByteArrayOutputStream out = new ByteArrayOutputStream();
//
//            // Landscape page
//            Document document = new Document(PageSize.A4.rotate());
//
//            PdfWriter.getInstance(document, out);
//            document.open();
//
//            // Find maximum number of columns
//            int maxColumn = 0;
//
//            for (Row row : sheet) {
//                if (row != null && row.getLastCellNum() > maxColumn) {
//                    maxColumn = row.getLastCellNum();
//                }
//            }
//
//            if (maxColumn == 0) {
//                document.add(new Phrase("The Excel file is empty."));
//                document.close();
//                return out.toByteArray();
//            }
//
//            // Create PDF table
//            PdfPTable pdfTable = new PdfPTable(maxColumn);
//            pdfTable.setWidthPercentage(100);
//
//            // ===========================
//            // Set Width Based On Header
//            // ===========================
//            Row headerRow = sheet.getRow(0);
//
//            float[] widths = new float[maxColumn];
//
//            for (int i = 0; i < maxColumn; i++) {
//
//                String header = "";
//
//                if (headerRow != null) {
//                    Cell cell = headerRow.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
//                    header = getCellValueAsString(cell);
//                }
//
//                // Width based on header text only
//                widths[i] = Math.max(header.length(), 8);
//
//                // Minimum width
//                if (widths[i] < 8)
//                    widths[i] = 8;
//
//                // Maximum width
//                if (widths[i] > 25)
//                    widths[i] = 25;
//            }
//
//            pdfTable.setWidths(widths);
//
//            // ===========================
//            // Fonts
//            // ===========================
//
//            Font headerFont = new Font(Font.HELVETICA, 11, Font.BOLD);
//            Font dataFont = new Font(Font.HELVETICA, 10);
//  
//            // ===========================
//            // Header Row
//            // ===========================
//            if (headerRow != null) {
//
//                for (int i = 0; i < maxColumn; i++) {
//
//                    Cell cell = headerRow.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
//
//                    PdfPCell pdfCell = new PdfPCell(
//                            new Phrase(getCellValueAsString(cell), headerFont));
//
//                    pdfCell.setHorizontalAlignment(Element.ALIGN_CENTER);
//                    pdfCell.setVerticalAlignment(Element.ALIGN_MIDDLE);
//                    pdfCell.setPadding(6);
//
//                    pdfTable.addCell(pdfCell);
//                }
//            }
//
//            // ===========================
//            // Data Rows
//            // ===========================
//            for (int r = 1; r <= sheet.getLastRowNum(); r++) {
//
//                Row row = sheet.getRow(r);
//
//                if (row == null)
//                    continue;
//
//                for (int c = 0; c < maxColumn; c++) {
//
//                    Cell cell = row.getCell(c, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
//
//                    PdfPCell pdfCell = new PdfPCell(
//                            new Phrase(getCellValueAsString(cell), dataFont));
//
//                    pdfCell.setPadding(5);
//                    pdfCell.setHorizontalAlignment(Element.ALIGN_LEFT);
//                    pdfCell.setVerticalAlignment(Element.ALIGN_MIDDLE);
//
//                    // Wrap long text
//                    pdfCell.setNoWrap(false);
//
//                    pdfTable.addCell(pdfCell);
//                }
//            }
//
//            document.add(pdfTable);
//            document.close();
//
//            return out.toByteArray();
//        }
		
		private final DataFormatter formatter = new DataFormatter();

	    private static final float PAGE_MARGIN = 20f;
	    private static final float HEADER_FONT_SIZE = 10f;
	    private static final float DATA_FONT_SIZE = 9f;
	    private static final float CELL_PADDING = 5f;
	    private static final float MIN_COLUMN_WIDTH = 8f;
	    private static final float MAX_COLUMN_WIDTH = 35f;

	    @Override
	    public byte[] convertExcelToPdf(InputStream excelStream) throws Exception {

	        try (Workbook workbook = WorkbookFactory.create(excelStream)) {

	            Sheet sheet = workbook.getSheetAt(0);

	            ByteArrayOutputStream out = new ByteArrayOutputStream();

	            Document document = new Document(PageSize.A3.rotate(), PAGE_MARGIN,
	                    PAGE_MARGIN, PAGE_MARGIN, PAGE_MARGIN);

	            PdfWriter.getInstance(document, out);

	            document.open();

	            int maxColumns = getMaxColumns(sheet);

	            if (maxColumns == 0) {
	                document.add(new Phrase("Excel file is empty."));
	                document.close();
	                return out.toByteArray();
	            }

	            PdfPTable table = new PdfPTable(maxColumns);

	            table.setWidthPercentage(100);

	            table.setWidths(calculateColumnWidths(sheet, maxColumns));

	            Font headerFont = FontFactory.getFont(
	                    FontFactory.HELVETICA_BOLD,
	                    HEADER_FONT_SIZE);

	            Font dataFont = FontFactory.getFont(
	                    FontFactory.HELVETICA,
	                    DATA_FONT_SIZE);

	            addHeaderRow(sheet, table, maxColumns, headerFont);

	            addDataRows(sheet, table, maxColumns, dataFont);

	            document.add(table);

	            document.close();

		        return out.toByteArray();

	        }
	    }

	    private int getMaxColumns(Sheet sheet) {

	        int max = 0;

	        for (Row row : sheet) {

	            if (row == null)
	                continue;

	            if (row.getLastCellNum() > max) {
	                max = row.getLastCellNum();
	            }
	        }

	        return max;
	    }

	    private void addHeaderRow(
	            Sheet sheet,
	            PdfPTable table,
	            int maxColumns,
	            Font headerFont) {

	        Row headerRow = sheet.getRow(0);

	        if (headerRow == null)
	            return;

	        for (int c = 0; c < maxColumns; c++) {

	            Cell cell = headerRow.getCell(c,
	                    Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);

	            PdfPCell pdfCell =
	                    new PdfPCell(new Phrase(getCellValue(cell), headerFont));

	            pdfCell.setHorizontalAlignment(getHorizontalAlignment(cell));
	            pdfCell.setVerticalAlignment(Element.ALIGN_MIDDLE);

	            pdfCell.setPadding(CELL_PADDING);

	            pdfCell.setBackgroundColor(new java.awt.Color(230, 230, 230));

	            pdfCell.setBorderWidth(1f);

	            pdfCell.setNoWrap(false);

	            table.addCell(pdfCell);
	        }
	    }

	    private void addDataRows(
	            Sheet sheet,
	            PdfPTable table,
	            int maxColumns,
	            Font dataFont) {

	        for (int r = 1; r <= sheet.getLastRowNum(); r++) {

	            Row row = sheet.getRow(r);

	            if (row == null)
	                continue;

	            for (int c = 0; c < maxColumns; c++) {

	                Cell cell = row.getCell(c,
	                        Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);

	                PdfPCell pdfCell =
	                        new PdfPCell(new Phrase(getCellValue(cell), dataFont));

	                pdfCell.setPadding(CELL_PADDING);

	                pdfCell.setHorizontalAlignment(
	                        getHorizontalAlignment(cell));

	                pdfCell.setVerticalAlignment(
	                        Element.ALIGN_MIDDLE);

	                pdfCell.setBorderWidth(0.8f);

	                pdfCell.setNoWrap(false);

	                table.addCell(pdfCell);
	            }
	        }
	    }

	    private float[] calculateColumnWidths(
	            Sheet sheet,
	            int maxColumns) throws DocumentException {

	        float[] widths = new float[maxColumns];

	        for (int c = 0; c < maxColumns; c++) {

	            int longest = 0;

	            for (Row row : sheet) {

	                if (row == null)
	                    continue;

	                Cell cell = row.getCell(c,
	                        Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);

	                String value = getCellValue(cell);

	                if (value != null) {

	                    value = value.replace("\n", " ");

	                    if (value.length() > longest) {
	                        longest = value.length();
	                    }
	                }
	            }

	            float excelWidth = sheet.getColumnWidth(c) / 256f;

	            float calculated =
	                    Math.max(excelWidth, longest * 0.75f);

	            if (calculated < MIN_COLUMN_WIDTH)
	                calculated = MIN_COLUMN_WIDTH;

	            if (calculated > MAX_COLUMN_WIDTH)
	                calculated = MAX_COLUMN_WIDTH;

	            widths[c] = calculated;
	        }

	        return widths;
	    }

	    private String getCellValue(Cell cell) {
 
	    	 if (cell == null)
	    	        return "";

	    	    String value = formatter.formatCellValue(cell);

	    	    if (value == null)
	    	        return "";

	    	    value = value.replace("\r", "");
	    	    value = value.replace("\n", " ");

	    	    return wrapText(value, 18);
	    }

	    private String wrapText(String text, int maxCharactersPerLine) {

	        if (text == null || text.isBlank())
	            return "";

	        StringBuilder result = new StringBuilder();

	        String[] words = text.trim().split("\\s+");

	        int currentLength = 0;

	        for (String word : words) {

	            if (currentLength == 0) {
	                result.append(word);
	                currentLength = word.length();
	            }
	            else if (currentLength + word.length() + 1 <= maxCharactersPerLine) {

	                result.append(" ").append(word);
	                currentLength += word.length() + 1;
	            }
	            else {

	                result.append("\n");
	                result.append(word);
	                currentLength = word.length();
	            }
	        }

	        return result.toString();
	    }
	    
	    private int getHorizontalAlignment(Cell cell) {

	        if (cell == null)
	            return Element.ALIGN_LEFT;

	        CellStyle style = cell.getCellStyle();

	        if (style == null)
	            return Element.ALIGN_LEFT;

	        HorizontalAlignment alignment =
	                style.getAlignment();

	        switch (alignment) {

	        case CENTER:
	        case CENTER_SELECTION:
	            return Element.ALIGN_CENTER;

	        case RIGHT:
	            return Element.ALIGN_RIGHT;

	        default:
	            return Element.ALIGN_LEFT;
	        }
    }

    private String getCellValueAsString(Cell cell) {

        if (cell == null)
            return "";

        switch (cell.getCellType()) {

        case STRING:
            return cell.getStringCellValue();

        case NUMERIC:
            if (DateUtil.isCellDateFormatted(cell)) {
                return cell.getDateCellValue().toString();
            }
            return String.valueOf(cell.getNumericCellValue());

        case BOOLEAN:
            return String.valueOf(cell.getBooleanCellValue());

        case FORMULA:
            try {
                return cell.getStringCellValue();
            } catch (Exception e) {
                try {
                    return String.valueOf(cell.getNumericCellValue());
                } catch (Exception ex) {
                    return cell.getCellFormula();
                }
            }

        case BLANK:
            return "";

        default:
            return "";
        }
    }
}