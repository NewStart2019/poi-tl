package com.deepoove.poi.util;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Calendar;
import java.util.Date;

public class ExcelUtils {

    /**
     * Copy from source line to target line, including content style and line height.
     *
     * @param sourceRow      {@link Row sourceRow} source row
     * @param targetRow      {@link Row workbook} target row
     * @param isIncludeStyle {@link Boolean isIncludeStyle} is include style
     */
    public static void copyRow(Row sourceRow, Row targetRow, boolean isIncludeStyle) {
        if (targetRow == null || sourceRow == null) {
            throw new NullPointerException("source row or target row is null");
        }
        Sheet sheet = targetRow.getSheet();
        if (isIncludeStyle) {
            targetRow.setHeight(sourceRow.getHeight());
            targetRow.setHeightInPoints(sourceRow.getHeightInPoints());
            targetRow.setRowStyle(sourceRow.getRowStyle());
        }

        for (int i = 0; i < sourceRow.getLastCellNum(); i++) {
            Cell oldCell = sourceRow.getCell(i);
            Cell newCell = targetRow.createCell(i);
            copyCell(oldCell, newCell, isIncludeStyle);
        }

        // 复制合并单元格: 到下一行
        copyMergedRegionFromTo(sheet, sourceRow.getRowNum(), targetRow.getRowNum());
    }

    /**
     * 将 sourceRowNum 行的合并单元格复制到 targetRowNum 行，
     * 并将 targetRowNum 及之后的所有合并区域整体下移一行。
     */
    public static void copyMergedRegionFromTo(Sheet sheet, int sourceRowNum, int targetRowNum) {
        for (CellRangeAddress region : sheet.getMergedRegions()) {
            // 构造对应在 targetRowNum 的新区域
            if (region.getFirstRow() == sourceRowNum && region.getLastRow() == sourceRowNum) {
                CellRangeAddress copiedRegion = new CellRangeAddress(
                    targetRowNum,
                    targetRowNum,
                    region.getFirstColumn(),
                    region.getLastColumn()
                );
                sheet.addMergedRegion(copiedRegion);
            }
        }
    }


    /**
     * Copy cell value, return directly if the <b>oldCell</b> or <b>newCell</b> is <b>empty</b>
     *
     * @param oldCell        {@link Cell oldCell} old cell
     * @param newCell        {@link Cell newCell} new cell
     * @param isIncludeStyle {@link Boolean isIncludeStyle} is include style
     */
    public static void copyCell(Cell oldCell, Cell newCell, boolean isIncludeStyle) {
        if (newCell == null || oldCell == null) {
            return;
        }
        if (isIncludeStyle) {
            newCell.setCellStyle(oldCell.getCellStyle());
        }
        switch (oldCell.getCellType()) {
            case STRING:
                newCell.setCellValue(oldCell.getStringCellValue());
                break;
            case NUMERIC:
                newCell.setCellValue(oldCell.getNumericCellValue());
                break;
            case BOOLEAN:
                newCell.setCellValue(oldCell.getBooleanCellValue());
                break;
            case FORMULA:
                newCell.setCellFormula(oldCell.getCellFormula());
                break;
            default:
                newCell.setCellValue("");
                break;
        }
    }

    /**
     * Set cell value, return directly if the <b>cell</b> or <b>value</b> is <b>empty</b>
     *
     * @param cell  {@link Cell cell}
     * @param value {@link Object value}
     */
    public static void setCellValue(Cell cell, Object value) {
        if (value == null || cell == null) {
            return;
        }
        if (value instanceof String) {
            cell.setCellValue((String) value);
        } else if (value instanceof Integer) {
            cell.setCellValue((Integer) value);
        } else if (value instanceof Double) {
            cell.setCellValue((Double) value);
        } else if (value instanceof Boolean) {
            cell.setCellValue((Boolean) value);
        } else if (value instanceof Date) {
            cell.setCellValue((Date) value);
        } else if (value instanceof Calendar) {
            cell.setCellValue((Calendar) value);
        } else if (value instanceof LocalDate) {
            cell.setCellValue((LocalDate) value);
        } else if (value instanceof LocalDateTime) {
            cell.setCellValue((LocalDateTime) value);
        } else if (value instanceof RichTextString) {
            cell.setCellValue((RichTextString) value);
        } else {
            cell.setCellValue(value.toString());
        }
    }

    /**
     * Insert a row at the specified position and move the following rows down
     */
    public static Row insertRowAndShiftBelow(Sheet sheet, int insertPosition) {
        int lastRowNum = sheet.getLastRowNum();
        if (insertPosition <= lastRowNum) {
            sheet.shiftRows(insertPosition, lastRowNum, 1, true, false);
        }
        return sheet.createRow(insertPosition);
    }


    public static void removeOneRowAndShift(Sheet sheet, int rowIndex) {
        if (rowIndex < 0 || rowIndex > sheet.getLastRowNum()) {
            return;
        }

        // delete row
        sheet.removeRow(sheet.getRow(rowIndex));
        // Move rowIndex+1 up to the last row as a whole by one row
        sheet.shiftRows(rowIndex + 1, sheet.getLastRowNum(), -1, true, true);
    }


    public static void removeOneRowAndShift(Sheet sheet, Row row) {
        removeOneRowAndShift(sheet, row.getRowNum());
    }
}
