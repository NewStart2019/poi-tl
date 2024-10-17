package com.deepoove.poi.util;

import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlObject;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.springframework.util.CollectionUtils;

import java.util.List;

/**
 * copy → clean → remove → find → set
 */
@SuppressWarnings("unused")
public class WordTableUtils {


    /**
     * Copy the content of the current line to the next line. If the next line is a newly added line,
     * then directly copy the entire XML of the current line to the next line. Otherwise, just copy
     * the content to the next line.
     *
     * @param currentLine      current line
     * @param nextLine         next line
     * @param templateRowIndex next line row index
     */
    public static XWPFTableRow copyLineContent(XWPFTableRow currentLine, XWPFTableRow nextLine, int templateRowIndex) {
        XWPFTable table = currentLine.getTable();
        if (org.apache.commons.collections4.CollectionUtils.isEmpty(nextLine.getTableCells())) {
            // 复制行
            XmlCursor sourceCursor = currentLine.getCtRow().newCursor();
            XmlObject object = sourceCursor.getObject();
            XmlObject targetXmlObject = nextLine.getCtRow().newCursor().getObject();
            targetXmlObject = targetXmlObject.set(object);
            nextLine = new XWPFTableRow((CTRow) targetXmlObject, table);
            setTableRow(table, nextLine, templateRowIndex);
        } else {
            List<XWPFTableCell> tableCells = currentLine.getTableCells();
            int nextCellSize = nextLine.getTableCells().size() - 1;
            for (int i = 0; i < tableCells.size(); i++) {
                XWPFTableCell currentCell = tableCells.get(i);
                boolean isNoHasCell = nextCellSize < i;
                XWPFTableCell nextCell = isNoHasCell ? nextLine.addNewTableCell() : nextLine.getCell(i);
                copyCellContent(currentCell, nextCell, isNoHasCell);
            }
            if (nextLine.getCtRow().isSetTrPr()) {
                nextLine.getCtRow().unsetTrPr();
                nextLine.getCtRow().setTrPr(currentLine.getCtRow().getTrPr());
            }
        }
        return nextLine;
    }

    /**
     * Copy the content of a cell that spans multiple columns, including its style. Since the data spanning columns
     * is only present in the first row, there's no need to clear the content of the target cells
     *
     * @param source source
     * @param target target
     */
    public static void copyCellContent(XWPFTableCell source, XWPFTableCell target, boolean isIncludeStyle) {
        List<XWPFParagraph> paragraphs = source.getParagraphs();
        CTPPr targetCtPPr = null;
        XWPFParagraph firstParagraph = null;
        if (org.apache.commons.collections4.CollectionUtils.isNotEmpty(target.getParagraphs())) {
            firstParagraph = target.getParagraphs().get(0);
            targetCtPPr = firstParagraph.getCTP().getPPr();
        }
        for (XWPFParagraph paragraph : source.getParagraphs()) {
            XWPFParagraph newParagraph = target.addParagraph();
            WordTableUtils.copyParagraphContent(paragraph, newParagraph);
            if (targetCtPPr != null) {
                // 复制段落样式
                newParagraph.getCTP().setPPr(targetCtPPr);
                newParagraph.setStyle(firstParagraph.getStyle());
            }
            if (isIncludeStyle) {
                newParagraph.getCTP().unsetPPr();
                newParagraph.getCTP().setPPr(paragraph.getCTP().getPPr());
            }
        }
        if (isIncludeStyle) {
            CTTcPr sourceTcPr = source.getCTTc().getTcPr();
            if (sourceTcPr != null) {
                target.getCTTc().setTcPr(source.getCTTc().getTcPr());
            }
        }
        if (firstParagraph != null) {
            WordTableUtils.removeParagraph(target, firstParagraph);
        }
    }

    /**
     * Copy the content of the spanned cells including the style. Since the spanned data only exists in the first row,
     * there is no need to clear the content of the target cells.
     *
     * @param source source
     * @param target target
     */
    public static void copyVerticallyCellContent(XWPFTableCell source, XWPFTableCell target) {
        removeAllParagraphs(target);
        for (XWPFParagraph paragraph : source.getParagraphs()) {
            XWPFParagraph newParagraph = target.addParagraph();
            copyParagraphContent(paragraph, newParagraph);
        }
        target.getCTTc().setTcPr(source.getCTTc().getTcPr());
    }

    public static void copyCellToNextRow(XWPFTable table, int rowIndex, int colIndex) {
        XWPFTableRow currentRow = table.getRow(rowIndex);
        XWPFTableCell currentCell = currentRow.getCell(colIndex);

        XWPFTableRow nextRow = table.getRow(rowIndex + 1);
        if (nextRow == null) {
            return;
        }

        XWPFTableCell newCell = nextRow.getCell(colIndex);
        if (newCell == null) {
            newCell = nextRow.createCell();
        }

        copyVerticallyCellContent(currentCell, newCell);
        cleanCellContent(currentCell);
    }

    public static void copyParagraphContent(XWPFParagraph source, XWPFParagraph target) {
        for (XWPFRun run : source.getRuns()) {
            XWPFRun newRun = target.createRun();
            CTRPr sourceCTRPr = run.getCTR().getRPr();
            CTRPr newCTRPr = CTRPr.Factory.newInstance();
            copyCTRPr(sourceCTRPr, newCTRPr);
            newRun.getCTR().setRPr(newCTRPr);
            newRun.setText(run.text());
        }
    }

    public static void copyCTRPr(CTRPr sourceCTRPr, CTRPr targetCTRPr) {
        // 验证 sourceCTRPr 和 targetCTRPr 是非空的
        if (sourceCTRPr == null || targetCTRPr == null) {
            throw new IllegalArgumentException("CTRPr objects cannot be null");
        }

        // 使用 XMLBeans 的 deepCopy 方法来复制属性
        XmlObject sourceXmlObject = sourceCTRPr.newCursor().getObject();
        XmlObject targetXmlObject = targetCTRPr.newCursor().getObject();
        // 深度复制源对象到目标对象
        targetXmlObject.set(sourceXmlObject);

        // 更新目标 CTRPr 对象
        targetCTRPr.set(targetXmlObject);
    }

    public static void cleanRowTextContent(XWPFTable table, int rowIndex) {
        cleanRowTextContent(table.getRow(rowIndex));
    }

    public static void cleanRowTextContent(XWPFTableRow templateRow) {
        List<XWPFTableCell> tableCells = templateRow.getTableCells();
        tableCells.forEach(cell -> {
            if (org.apache.commons.collections4.CollectionUtils.isNotEmpty(cell.getParagraphs())) {
                cell.getParagraphs().forEach(WordTableUtils::removeAllRun);
            }
        });
    }

    public static void cleanCellContent(XWPFTableCell cell) {
        cell.removeParagraph(0);
    }

    public static void removeAllParagraphs(XWPFTableCell cell) {
        List<XWPFParagraph> paragraphs = cell.getParagraphs();
        int size = paragraphs.size();
        for (int i = size - 1; i >= 0; i--) {
            cell.removeParagraph(i);
        }
    }

    public static void removeParagraph(XWPFTableCell cell, XWPFParagraph paragraph) {
        if (!CollectionUtils.isEmpty(cell.getParagraphs())) {
            cell.removeParagraph(cell.getParagraphs().indexOf(paragraph));
        }
    }

    public static void removeRun(XWPFParagraph paragraph, XWPFRun run) {
        if (!CollectionUtils.isEmpty(paragraph.getRuns())) {
            paragraph.removeRun(paragraph.getRuns().indexOf(run));
        }
    }

    public static void removeAllRun(XWPFParagraph paragraph) {
        if (paragraph != null && org.apache.commons.collections4.CollectionUtils.isNotEmpty(paragraph.getRuns())) {
            for (int i = paragraph.getRuns().size() - 1; i >= 0; i--) {
                paragraph.removeRun(i);
            }
        }
    }

    /**
     * Get the number of columns spanned in the table (Issue: If the columns are misaligned, the handling method has problems)
     *
     * @param table    XWPFTable
     * @param startRow start row index
     * @param colIndex col index
     * @return span col number，0 indicates no cross row
     */
    public static int findMergedRows(XWPFTable table, int startRow, int colIndex) {
        int i = startRow + 1;
        int size = table.getRows().size();
        for (; i <= size; i++) {
            if (table.getRow(i).getTableCells().get(colIndex) == null) {
                break;
            }
            XWPFTableCell xwpfTableCell = table.getRow(i).getTableCells().get(colIndex);
            if (xwpfTableCell.getCTTc() == null) {
                break;
            }
            CTTc ctTc = xwpfTableCell.getCTTc();
            if (ctTc.getTcPr() == null) {
                break;
            }
            if (ctTc.getTcPr().getVMerge() == null) {
                break;
            }
            if (ctTc.getTcPr().getVMerge().getVal().intValue() != 1) {
                break;
            }
        }
        return i - startRow;
    }

    public static int findRowIndex(XWPFTableCell tagCell) {
        XWPFTableRow tagRow = tagCell.getTableRow();
        return findRowIndex(tagRow);
    }

    public static int findRowIndex(XWPFTableRow row) {
        List<XWPFTableRow> rows = row.getTable().getRows();
        return rows.indexOf(row);
    }

    /**
     * Retrieve the spanned row data, where restart=2 indicates the start of a span.
     * continue=1 signifies the continuation of the spanned data, and the spanning ends when there is no more span information.
     *
     * @param cell
     * @return Integer | null则表示没有跨行
     */
    public static Integer findVMerge(XWPFTableCell cell) {
        // Get cell properties
        CTTcPr tcPr = cell.getCTTc().getTcPr();
        if (tcPr != null) {
            // Get vertical merge properties
            CTVMerge vMerge = tcPr.getVMerge();
            if (vMerge != null) {
                return vMerge.getVal().intValue();
            }
        }
        return null;
    }

    /**
     * Get the maximum actual number of columns in the table (including all spanned rows)
     *
     * @param table XWPFTable table
     * @return int
     */
    public static int findMaxColIncludeSpanCol(XWPFTable table) {
        int max = -1;
        for (XWPFTableRow row : table.getRows()) {
            List<XWPFTableCell> tableCells = row.getTableCells();
            int temp = 0;
            for (XWPFTableCell cell : tableCells) {
                if (cell.getCTTc() == null) {
                    temp += 1;
                    continue;
                }
                CTTc ctTc = cell.getCTTc();
                if (ctTc.getTcPr() == null) {
                    temp += 1;
                    continue;
                }
                if (ctTc.getTcPr().getGridSpan() == null) {
                    temp += 1;
                } else {
                    temp += ctTc.getTcPr().getGridSpan().getVal().intValue();
                }
            }
            if (temp > max) {
                max = temp;
            }
        }
        return max;
    }


    /**
     * Get the maximum number of columns in the table
     *
     * @param table XWPFTable table
     * @return int
     */
    public static int findMaxSpanColCount(XWPFTable table) {
        int maxColCount = 0;
        List<XWPFTableRow> rows = table.getRows();
        for (XWPFTableRow row : rows) {
            int colCount = row.getTableCells().size();
            if (colCount > maxColCount) {
                maxColCount = colCount;
            }
        }
        return maxColCount;
    }

    @SuppressWarnings("unchecked")
    public static void setTableRow(XWPFTable table, XWPFTableRow row, int pos) {
        List<XWPFTableRow> rows = (List<XWPFTableRow>) ReflectionUtils.getValue("tableRows", table);
        rows.set(pos, row);
        table.getCTTbl().setTrArray(pos, row.getCtRow());
    }

}
