package util;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * Utility class that scans all sheets in an Excel workbook for the Vietnamese
 * customs declaration fields "Mã số hàng hóa" and "Mô tả hàng hóa", extracts
 * their corresponding values, and writes the results to a new sheet named "Result".
 */
public class ExcelProcessor {

    private static final String LABEL_MA_SO = "Mã số hàng hóa";
    private static final String LABEL_MO_TA = "Mô tả hàng hóa";
    private static final String RESULT_SHEET = "Result";

    /**
     * Reads the Excel workbook from the given input stream, processes all
     * sheets, and writes the enriched workbook (with a "Result" sheet) to the
     * output stream.
     *
     * @param inputStream  the Excel file input stream (.xls or .xlsx)
     * @param outputStream the output stream to write the modified workbook
     * @param fileName     the original file name – used to decide the format
     * @throws IOException if an I/O error occurs
     */
    public static void process(InputStream inputStream, OutputStream outputStream,
                               String fileName) throws IOException {
        Workbook workbook = openWorkbook(inputStream, fileName);
        try {
            List<String[]> records = extractRecords(workbook);
            writeResultSheet(workbook, records);
            workbook.write(outputStream);
        } finally {
            workbook.close();
        }
    }

    // -----------------------------------------------------------------------
    // Private helpers
    // -----------------------------------------------------------------------

    private static Workbook openWorkbook(InputStream is, String fileName) throws IOException {
        try {
            if (fileName != null && fileName.toLowerCase().endsWith(".xls")) {
                return new HSSFWorkbook(is);
            }
            return new XSSFWorkbook(is);
        } catch (Exception e) {
            throw new IOException(
                "Không thể mở tệp Excel. Tệp có thể bị hỏng hoặc không đúng định dạng.", e);
        }
    }

    /**
     * Iterates over every sheet (except a pre-existing "Result" sheet) and
     * searches for cells whose trimmed text equals one of the two target
     * labels.  When found, the value is taken from the cell immediately to
     * the right.
     */
    private static List<String[]> extractRecords(Workbook workbook) {
        List<String[]> records = new ArrayList<>();

        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            Sheet sheet = workbook.getSheetAt(i);
            String sheetName = sheet.getSheetName();

            // Skip the result sheet if it already exists
            if (RESULT_SHEET.equalsIgnoreCase(sheetName)) {
                continue;
            }

            String maSo = null;
            String moTa = null;

            for (Row row : sheet) {
                for (Cell cell : row) {
                    String cellText = getCellText(cell);
                    if (cellText == null) {
                        continue;
                    }

                    if (LABEL_MA_SO.equalsIgnoreCase(cellText.trim())) {
                        maSo = getAdjacentValue(cell);
                    } else if (LABEL_MO_TA.equalsIgnoreCase(cellText.trim())) {
                        moTa = getAdjacentValue(cell);
                    }
                }
            }

            // Only add a row when at least one of the fields was found and
            // its value is non-empty (requirement 4 & 8).
            boolean hasMaSo = maSo != null && !maSo.isEmpty();
            boolean hasMoTa = moTa != null && !moTa.isEmpty();
            if (hasMaSo || hasMoTa) {
                records.add(new String[]{sheetName, maSo != null ? maSo : "", moTa != null ? moTa : ""});
            }
        }

        return records;
    }

    /**
     * Returns the value of the cell immediately to the right of the given
     * label cell.  Returns {@code null} if the adjacent cell does not exist
     * or is blank.
     */
    private static String getAdjacentValue(Cell labelCell) {
        Row row = labelCell.getRow();
        int nextCol = labelCell.getColumnIndex() + 1;
        Cell valueCell = row.getCell(nextCol);
        if (valueCell == null) {
            return null;
        }
        String value = getCellText(valueCell);
        return (value != null && !value.trim().isEmpty()) ? value.trim() : null;
    }

    /** Returns the string representation of a cell, or {@code null} for blank cells. */
    private static String getCellText(Cell cell) {
        if (cell == null) {
            return null;
        }
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                }
                double numVal = cell.getNumericCellValue();
                // Return as integer string when the value is a whole number
                if (numVal == Math.floor(numVal) && !Double.isInfinite(numVal)) {
                    return String.valueOf((long) numVal);
                }
                return String.valueOf(numVal);
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                try {
                    return cell.getStringCellValue();
                } catch (IllegalStateException e) {
                    return String.valueOf(cell.getNumericCellValue());
                }
            case BLANK:
            default:
                return null;
        }
    }

    /**
     * Removes any existing "Result" sheet from the workbook and creates a new
     * one populated with the extracted records.
     */
    private static void writeResultSheet(Workbook workbook, List<String[]> records) {
        // Remove existing result sheet if present
        int existingIndex = workbook.getSheetIndex(RESULT_SHEET);
        if (existingIndex != -1) {
            workbook.removeSheetAt(existingIndex);
        }

        Sheet resultSheet = workbook.createSheet(RESULT_SHEET);

        // Header row
        Row header = resultSheet.createRow(0);
        header.createCell(0).setCellValue("Sheet Name");
        header.createCell(1).setCellValue(LABEL_MA_SO);
        header.createCell(2).setCellValue(LABEL_MO_TA);

        // Data rows
        int rowIndex = 1;
        for (String[] record : records) {
            Row dataRow = resultSheet.createRow(rowIndex++);
            dataRow.createCell(0).setCellValue(record[0]);
            dataRow.createCell(1).setCellValue(record[1]);
            dataRow.createCell(2).setCellValue(record[2]);
        }

        // Auto-size columns for readability
        for (int col = 0; col < 3; col++) {
            resultSheet.autoSizeColumn(col);
        }
    }
}
