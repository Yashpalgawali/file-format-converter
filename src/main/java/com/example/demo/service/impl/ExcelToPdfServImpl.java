package com.example.demo.service.impl;

import java.io.ByteArrayOutputStream;
import java.io.InputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import com.example.demo.service.IExcelToPdfService;
import com.lowagie.text.Document;
import com.lowagie.text.Element;
import com.lowagie.text.Font;
import com.lowagie.text.PageSize;
import com.lowagie.text.Phrase;
import com.lowagie.text.pdf.PdfPCell;
import com.lowagie.text.pdf.PdfPTable;
import com.lowagie.text.pdf.PdfWriter;

@Service("exceltopdfserv")
public class ExcelToPdfServImpl implements IExcelToPdfService {
	@Override
    public byte[] convertExcelToPdf(InputStream excelStream) throws Exception {

        try (Workbook workbook = new XSSFWorkbook(excelStream)) {

            Sheet sheet = workbook.getSheetAt(0);

            ByteArrayOutputStream out = new ByteArrayOutputStream();

            // Landscape page
            Document document = new Document(PageSize.A4.rotate());

            PdfWriter.getInstance(document, out);
            document.open();

            // Find maximum number of columns
            int maxColumn = 0;

            for (Row row : sheet) {
                if (row != null && row.getLastCellNum() > maxColumn) {
                    maxColumn = row.getLastCellNum();
                }
            }

            if (maxColumn == 0) {
                document.add(new Phrase("The Excel file is empty."));
                document.close();
                return out.toByteArray();
            }

            // Create PDF table
            PdfPTable pdfTable = new PdfPTable(maxColumn);
            pdfTable.setWidthPercentage(100);

            // ===========================
            // Set Width Based On Header
            // ===========================
            Row headerRow = sheet.getRow(0);

            float[] widths = new float[maxColumn];

            for (int i = 0; i < maxColumn; i++) {

                String header = "";

                if (headerRow != null) {
                    Cell cell = headerRow.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    header = getCellValueAsString(cell);
                }

                // Width based on header text only
                widths[i] = Math.max(header.length(), 8);

                // Minimum width
                if (widths[i] < 8)
                    widths[i] = 8;

                // Maximum width
                if (widths[i] > 25)
                    widths[i] = 25;
            }

            pdfTable.setWidths(widths);

            // ===========================
            // Fonts
            // ===========================

            Font headerFont = new Font(Font.HELVETICA, 11, Font.BOLD);
            Font dataFont = new Font(Font.HELVETICA, 10);
  
            // ===========================
            // Header Row
            // ===========================
            if (headerRow != null) {

                for (int i = 0; i < maxColumn; i++) {

                    Cell cell = headerRow.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);

                    PdfPCell pdfCell = new PdfPCell(
                            new Phrase(getCellValueAsString(cell), headerFont));

                    pdfCell.setHorizontalAlignment(Element.ALIGN_CENTER);
                    pdfCell.setVerticalAlignment(Element.ALIGN_MIDDLE);
                    pdfCell.setPadding(6);

                    pdfTable.addCell(pdfCell);
                }
            }

            // ===========================
            // Data Rows
            // ===========================
            for (int r = 1; r <= sheet.getLastRowNum(); r++) {

                Row row = sheet.getRow(r);

                if (row == null)
                    continue;

                for (int c = 0; c < maxColumn; c++) {

                    Cell cell = row.getCell(c, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);

                    PdfPCell pdfCell = new PdfPCell(
                            new Phrase(getCellValueAsString(cell), dataFont));

                    pdfCell.setPadding(5);
                    pdfCell.setHorizontalAlignment(Element.ALIGN_LEFT);
                    pdfCell.setVerticalAlignment(Element.ALIGN_MIDDLE);

                    // Wrap long text
                    pdfCell.setNoWrap(false);

                    pdfTable.addCell(pdfCell);
                }
            }

            document.add(pdfTable);
            document.close();

            return out.toByteArray();
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