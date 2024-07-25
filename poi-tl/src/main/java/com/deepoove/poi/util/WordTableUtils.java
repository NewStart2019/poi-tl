package com.deepoove.poi.util;

import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTVMerge;
import org.springframework.util.CollectionUtils;

import java.util.Collections;
import java.util.List;

public class WordTableUtils {

    /**
     * 获取表格的最大列数
     *
     * @param table XWPFTable 表格
     * @return 最大列数
     */
    public static int getMaxColCount(XWPFTable table) {
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

    /**
     * 获取表格的最大实际列数（包括所有跨行）
     *
     * @param table XWPFTable 表格
     * @return 最大列数
     */
    public static int getMaxActualCol(XWPFTable table) {
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
                    continue;
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
     * 获取跨行数据，restart=2 表示跨行的开始
     * continue=1是跨行数据的持续，知道跨行信息不存在则结束跨行
     *
     * @param cell 单元格
     * @return null则表示没有跨行
     */
    public static Integer getVMerge(XWPFTableCell cell) {
        // 获取单元格属性
        CTTcPr tcPr = cell.getCTTc().getTcPr();
        if (tcPr != null) {
            // 获取垂直合并属性
            CTVMerge vMerge = tcPr.getVMerge();
            if (vMerge != null) {
                return vMerge.getVal().intValue();
            }
        }
        return null;
    }

    /**
     * 将当前单元格移动到下一行同一列
     *
     * @param table    XWPFTable 表格
     * @param rowIndex 当前行索引
     * @param colIndex 当前列索引
     */
    public static void moveCellToNextRow(XWPFTable table, int rowIndex, int colIndex) {
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
        clearCellContent(currentCell);
    }

    /**
     * 复制跨列的单元格内容包括样式，由于跨列的数据只在第一一行有，所以不需要清除目标单元格的内容
     *
     * @param source source
     * @param target target
     */
    public static void copyVerticallyCellContent(XWPFTableCell source, XWPFTableCell target) {
        removeAllParagraphs(target);
        for (XWPFParagraph paragraph : source.getParagraphs()) {
            XWPFParagraph newParagraph = target.addParagraph();
            copyParagraph(paragraph, newParagraph);
        }
        target.getCTTc().setTcPr(source.getCTTc().getTcPr());
    }

    // 复制段落内容
    public static void copyParagraph(XWPFParagraph source, XWPFParagraph target) {
        for (XWPFRun run : source.getRuns()) {
            XWPFRun newRun = target.createRun();
            newRun.getCTR().setRPr(run.getCTR().getRPr());
            newRun.setText(run.text());
        }
    }

    // 清空单元格内容
    public static void clearCellContent(XWPFTableCell cell) {
        cell.removeParagraph(0);
    }

    /**
     * 获取表格的跨列数（问题：如果列错位了处理方式就有问题）
     *
     * @param table    表格
     * @param startRow 开始行
     * @param colIndex 查找列
     * @return 跨行数目，如果为0表示没有跨行
     */
    public static int getMergedRows(XWPFTable table, int startRow, int colIndex) {
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

    /**
     * 删除单元格中的所有段落
     *
     * @param cell XWPFTableCell 单元格
     */
    public static void removeAllParagraphs(XWPFTableCell cell) {
        List<XWPFParagraph> paragraphs = cell.getParagraphs();
        int size = paragraphs.size();
        for (int i = size - 1; i >= 0; i--) {
            cell.removeParagraph(i);
        }
    }

    /**
     * 移除单元格段落
     *
     * @param cell      {@link XWPFTableCell cell}
     * @param paragraph {@link XWPFParagraph paragraph}
     */
    public static void removeParagraph(XWPFTableCell cell, XWPFParagraph paragraph) {
        if (!CollectionUtils.isEmpty(cell.getParagraphs())) {
            cell.removeParagraph(cell.getParagraphs().indexOf(paragraph));
        }
    }

    /**
     * 移除段落指定的run
     *
     * @param paragraph {@link XWPFParagraph paragraph}
     * @param run       {@link XWPFRun run}
     */
    public static void removeRun(XWPFParagraph paragraph, XWPFRun run) {
        if (!CollectionUtils.isEmpty(paragraph.getRuns())) {
            paragraph.removeRun(paragraph.getRuns().indexOf(run));
        }
    }

}
