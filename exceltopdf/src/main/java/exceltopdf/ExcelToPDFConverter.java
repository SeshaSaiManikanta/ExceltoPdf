package exceltopdf;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.List;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.itextpdf.text.*;
import com.itextpdf.text.Font;
import com.itextpdf.text.pdf.*;

public class ExcelToPDFConverter {

    public static void main(String[] args) {
        System.out.println("Excel to PDF conversion starting...");

        String inputDir = "S:\\Softwares\\input_excel";
        String outputDir = "S:\\Softwares\\output_pdf";

        File inputFolder = new File(inputDir);
        File[] files = inputFolder.listFiles((dir, name) -> name.endsWith(".xlsx"));

        if (files != null) {
            for (File excelFile : files) {
                File pdfFile = new File(outputDir, excelFile.getName().replace(".xlsx", ".pdf"));
                convertExcelToPDF(excelFile, pdfFile);
            }
        }

        System.out.println("Conversion completed.");
    }

 public static void convertExcelToPDF(File excelFile, File pdfFile) {
    try (FileInputStream fis = new FileInputStream(excelFile);
         Workbook workbook = WorkbookFactory.create(fis);
         FileOutputStream fos = new FileOutputStream(pdfFile)) {

        Document document = new Document(PageSize.A4.rotate(), 20, 20, 50, 40); // landscape
        PdfWriter writer = PdfWriter.getInstance(document, fos);

        // Add footer with page numbers
        writer.setPageEvent(new PdfPageEventHelper() {
            @Override
            public void onEndPage(PdfWriter writer, Document document) {
                ColumnText.showTextAligned(writer.getDirectContent(),
                        Element.ALIGN_CENTER,
                        new Phrase(String.format("Page %d", writer.getPageNumber())),
                        (document.right() + document.left()) / 2,
                        document.bottom() - 10, 0);
            }
        });

        document.open();

        // Title
        Font titleFont = FontFactory.getFont(FontFactory.HELVETICA_BOLD, 16, BaseColor.BLACK);
        Paragraph title = new Paragraph("Lekker India Working Hours - April", titleFont);
        title.setAlignment(Element.ALIGN_CENTER);
        title.setSpacingAfter(15);
        document.add(title);

        Sheet sheet = workbook.getSheetAt(0);
        if (sheet.getPhysicalNumberOfRows() == 0) {
            System.out.println("Sheet is empty, skipping PDF creation.");
            return;
        }

        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();

        // Get max number of columns
        int numCols = 0;
        for (Row row : sheet) {
            if (row.getLastCellNum() > numCols) {
                numCols = row.getLastCellNum();
            }
        }

        PdfPTable table = new PdfPTable(numCols);
        table.setWidthPercentage(100);
        Font tableFont = FontFactory.getFont(FontFactory.HELVETICA, 10);
        Font headerFont = FontFactory.getFont(FontFactory.HELVETICA_BOLD, 11, BaseColor.WHITE);

        // Background for header
        BaseColor headerColor = new BaseColor(0, 121, 182); // blue

        List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();
        boolean isHeader = true;
        
        Set<String> mergedCellPositions = new HashSet<>();
        for (CellRangeAddress region : mergedRegions) {
            for (int r = region.getFirstRow(); r <= region.getLastRow(); r++) {
                for (int c = region.getFirstColumn(); c <= region.getLastColumn(); c++) {
                    if (!(r == region.getFirstRow() && c == region.getFirstColumn())) {
                        mergedCellPositions.add(r + "-" + c);
                    }
                }
            }
        }


        for (Row row : sheet) {
            int colIndex = 0;
            if (mergedCellPositions.contains(row.getRowNum() + "-" + colIndex)) {
                colIndex++;
                continue;
            }


            while (colIndex < numCols) {
                Cell cell = row.getCell(colIndex);
                String text = getCellText(cell, evaluator);

                PdfPCell pdfCell = new PdfPCell(new Phrase(text, isHeader ? headerFont : tableFont));
                int colspan = 1;

                for (CellRangeAddress region : mergedRegions) {
                    if (region.getFirstRow() == row.getRowNum() && region.getFirstColumn() == colIndex) {
                        colspan = region.getLastColumn() - region.getFirstColumn() + 1;
                        break;
                    }
                }

                pdfCell.setColspan(colspan);
                pdfCell.setHorizontalAlignment(Element.ALIGN_CENTER);
                pdfCell.setVerticalAlignment(Element.ALIGN_MIDDLE);
                pdfCell.setPadding(5);

                if (isHeader) {
                    pdfCell.setBackgroundColor(headerColor);
                } else {
                    pdfCell.setBackgroundColor(row.getRowNum() % 2 == 0 ? BaseColor.LIGHT_GRAY : BaseColor.WHITE);
                }

                table.addCell(pdfCell);
                colIndex += colspan;
            }

            if (isHeader) isHeader = false;
        }

        document.add(table);
        document.close();
        System.out.println("PDF created: " + pdfFile.getAbsolutePath());

    } catch (Exception e) {
        e.printStackTrace();
    }
}



//working fine except Total hours & Total cost formulas cells
/*
 * private static String getCellText(Cell cell, FormulaEvaluator evaluator) { if
 * (cell == null) return "";
 * 
 * try { CellType cellType = cell.getCellType();
 * 
 * if (cellType == CellType.FORMULA) { CellValue evaluatedValue =
 * evaluator.evaluate(cell); //if (evaluatedValue == null) return "[Error]"; if
 * (evaluatedValue == null || evaluatedValue.getCellType() == CellType.ERROR) {
 * // Fallback to cached result CellType cachedType =
 * cell.getCachedFormulaResultType(); switch (cachedType) { case NUMERIC: if
 * (DateUtil.isCellDateFormatted(cell)) { Date date = cell.getDateCellValue();
 * return new SimpleDateFormat("h:mm a").format(date); } double fallbackVal =
 * cell.getNumericCellValue(); return String.format("%.2f", fallbackVal); case
 * STRING: return cell.getStringCellValue().trim(); case BOOLEAN: return
 * Boolean.toString(cell.getBooleanCellValue()); default: return "[Error]"; } }
 * switch (evaluatedValue.getCellType()) {
 * 
 * 
 * 
 * case STRING: return evaluatedValue.getStringValue().trim();
 * 
 * case NUMERIC: double value = evaluatedValue.getNumberValue();
 * 
 * // Date/time handling if (DateUtil.isValidExcelDate(value) &&
 * DateUtil.isCellDateFormatted(cell)) { String format =
 * cell.getCellStyle().getDataFormatString().toLowerCase(); Date date =
 * DateUtil.getJavaDate(value); if (format.contains("h") &&
 * !format.contains("d")) { return new SimpleDateFormat("h:mm a").format(date);
 * // Time only } else { return new SimpleDateFormat("MM/dd/yyyy").format(date);
 * } }
 * 
 * // "Hours Per Day" detection Row headerRow = cell.getSheet().getRow(0); if
 * (headerRow != null) { Cell headerCell =
 * headerRow.getCell(cell.getColumnIndex()); if (headerCell != null &&
 * "Hours Per Day".equalsIgnoreCase(headerCell.getStringCellValue().trim())) {
 * int totalMinutes = (int) Math.round(value * 24 * 60); int hours =
 * totalMinutes / 60; int minutes = totalMinutes % 60; return
 * String.format("%d:%02d", hours, minutes); } }
 * 
 * return String.format("%.2f", value);
 * 
 * case BOOLEAN: return String.valueOf(evaluatedValue.getBooleanValue());
 * 
 * case ERROR: return "[Error]";
 * 
 * default: return "[Unknown]"; }
 * 
 * } else if (cellType == CellType.NUMERIC) { double value =
 * cell.getNumericCellValue();
 * 
 * if (DateUtil.isCellDateFormatted(cell)) { String format =
 * cell.getCellStyle().getDataFormatString().toLowerCase(); Date date =
 * cell.getDateCellValue(); if (format.contains("h") && !format.contains("d")) {
 * return new SimpleDateFormat("h:mm a").format(date); } else { return new
 * SimpleDateFormat("MM/dd/yyyy").format(date); } }
 * 
 * // Check for hours formatting Row headerRow = cell.getSheet().getRow(0); if
 * (headerRow != null) { Cell headerCell =
 * headerRow.getCell(cell.getColumnIndex()); if (headerCell != null &&
 * "Hours Per Day".equalsIgnoreCase(headerCell.getStringCellValue().trim())) {
 * int totalMinutes = (int) Math.round(value * 24 * 60); int hours =
 * totalMinutes / 60; int minutes = totalMinutes % 60; return
 * String.format("%d:%02d", hours, minutes); } }
 * 
 * return String.format("%.2f", value);
 * 
 * } else if (cellType == CellType.STRING) { return
 * cell.getStringCellValue().trim();
 * 
 * } else if (cellType == CellType.BOOLEAN) { return
 * Boolean.toString(cell.getBooleanCellValue());
 * 
 * } else if (cellType == CellType.BLANK) { return "";
 * 
 * } else { return cell.toString(); }
 * 
 * } catch (Exception ex) { return "[Error]"; } }
 */
 
 private static String getCellText(Cell cell, FormulaEvaluator evaluator) {
	    if (cell == null) return "";

	    try {
	        CellType cellType = cell.getCellType();

	        if (cellType == CellType.FORMULA) {
	            CellValue evaluatedValue = evaluator.evaluate(cell);

	            // Check if formula evaluation failed or resulted in error
	            if (evaluatedValue == null || evaluatedValue.getCellType() == CellType.ERROR) {

	                // 🔁 Fallback to cached formula result type if formula evaluation fails
	                CellType cachedType = cell.getCachedFormulaResultType();

	                // This fallback is important especially for formulas like Total Hours/Total Cost
	                switch (cachedType) {
	                    case NUMERIC:
	                        if (DateUtil.isCellDateFormatted(cell)) {
	                            Date date = cell.getDateCellValue();
	                            return new SimpleDateFormat("h:mm a").format(date);
	                        }

	                        // 💡 Use cached numeric value as fallback
	                        double fallbackVal = cell.getNumericCellValue();

	                        // Handle same Hours Per Day fallback logic for formulas too
	                        Row headerRow = cell.getSheet().getRow(0);
	                        if (headerRow != null) {
	                            Cell headerCell = headerRow.getCell(cell.getColumnIndex());
	                            if (headerCell != null && "Hours Per Day".equalsIgnoreCase(headerCell.getStringCellValue().trim())) {
	                                int totalMinutes = (int) Math.round(fallbackVal * 24 * 60);
	                                int hours = totalMinutes / 60;
	                                int minutes = totalMinutes % 60;
	                                return String.format("%d:%02d", hours, minutes);
	                            }
	                        }

	                        return String.format("%.2f", fallbackVal);

	                    case STRING:
	                        return cell.getStringCellValue().trim();

	                    case BOOLEAN:
	                        return Boolean.toString(cell.getBooleanCellValue());

	                    default:
	                        return "[Error]";
	                }
	            }

	            // ✅ Evaluation successful, now handle evaluated value types
	            switch (evaluatedValue.getCellType()) {

	                case STRING:
	                    return evaluatedValue.getStringValue().trim();

	                case NUMERIC:
	                    double value = evaluatedValue.getNumberValue();

	                    // Handle date/time formats if present
	                    if (DateUtil.isValidExcelDate(value) && DateUtil.isCellDateFormatted(cell)) {
	                        String format = cell.getCellStyle().getDataFormatString().toLowerCase();
	                        Date date = DateUtil.getJavaDate(value);
	                        if (format.contains("h") && !format.contains("d")) {
	                            return new SimpleDateFormat("h:mm a").format(date); // Time only
	                        } else {
	                            return new SimpleDateFormat("MM/dd/yyyy").format(date);
	                        }
	                    }

	                    // Hours Per Day format logic for evaluated numeric
	                    Row headerRow = cell.getSheet().getRow(0);
	                    if (headerRow != null) {
	                        Cell headerCell = headerRow.getCell(cell.getColumnIndex());
	                        if (headerCell != null && "Hours Per Day".equalsIgnoreCase(headerCell.getStringCellValue().trim())) {
	                            int totalMinutes = (int) Math.round(value * 24 * 60);
	                            int hours = totalMinutes / 60;
	                            int minutes = totalMinutes % 60;
	                            return String.format("%d:%02d", hours, minutes);
	                        }
	                    }

	                    return String.format("%.2f", value);

	                case BOOLEAN:
	                    return String.valueOf(evaluatedValue.getBooleanValue());

	                case ERROR:
	                    return "[Error]";

	                default:
	                    return "[Unknown]";
	            }

	        } else if (cellType == CellType.NUMERIC) {
	            double value = cell.getNumericCellValue();

	            if (DateUtil.isCellDateFormatted(cell)) {
	                String format = cell.getCellStyle().getDataFormatString().toLowerCase();
	                Date date = cell.getDateCellValue();
	                if (format.contains("h") && !format.contains("d")) {
	                    return new SimpleDateFormat("h:mm a").format(date);
	                } else {
	                    return new SimpleDateFormat("MM/dd/yyyy").format(date);
	                }
	            }

	            // Same "Hours Per Day" logic for plain numeric cells
	            Row headerRow = cell.getSheet().getRow(0);
	            if (headerRow != null) {
	                Cell headerCell = headerRow.getCell(cell.getColumnIndex());
	                if (headerCell != null && "Hours Per Day".equalsIgnoreCase(headerCell.getStringCellValue().trim())) {
	                    int totalMinutes = (int) Math.round(value * 24 * 60);
	                    int hours = totalMinutes / 60;
	                    int minutes = totalMinutes % 60;
	                    return String.format("%d:%02d", hours, minutes);
	                }
	            }

	            return String.format("%.2f", value);

	        } else if (cellType == CellType.STRING) {
	            return cell.getStringCellValue().trim();

	        } else if (cellType == CellType.BOOLEAN) {
	            return Boolean.toString(cell.getBooleanCellValue());

	        } else if (cellType == CellType.BLANK) {
	            return "";

	        } else {
	            return cell.toString();
	        }

	    } catch (Exception ex) {
	        return "[Error]";
	    }
	}



    private static int getMaxColumns(Sheet sheet) {
        int max = 0;
        for (Row row : sheet) {
            max = Math.max(max, row.getLastCellNum());
        }
        return max;
    }

    // Footer class
    static class HeaderFooterPageEvent extends PdfPageEventHelper {
        Font footerFont = FontFactory.getFont(FontFactory.HELVETICA, 9, Font.ITALIC, BaseColor.GRAY);

        public void onEndPage(PdfWriter writer, Document document) {
            PdfContentByte cb = writer.getDirectContent();
            Phrase footer = new Phrase("Page " + document.getPageNumber(), footerFont);
            ColumnText.showTextAligned(cb, Element.ALIGN_CENTER,
                    footer, (document.right() + document.left()) / 2, document.bottom() - 10, 0);
        }
    }
}
