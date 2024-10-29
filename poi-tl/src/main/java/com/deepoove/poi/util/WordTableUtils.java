package com.deepoove.poi.util;

import com.deepoove.poi.xwpf.Page;
import com.deepoove.poi.xwpf.XWPFStructuredDocumentTagContent;
import com.deepoove.poi.xwpf.XWPFTextboxContent;
import com.sun.istack.internal.NotNull;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlObject;
import org.openxmlformats.schemas.drawingml.x2006.picture.CTPicture;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.util.CollectionUtils;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.util.List;

/**
 * copy → clean → remove → find → set → merge
 */
@SuppressWarnings("unused")
public class WordTableUtils {

    private static final Logger logger = LoggerFactory.getLogger(WordTableUtils.class);

    public static XWPFTable copyTable(XWPFDocument doc, XWPFTable sourceTable) {
        return copyTable(doc, sourceTable, false);
    }

    /**
     * <p>Copy a new table from the original document and place it after the original table.</p>
     * <p>If isTail is True, the new table will be placed at the end of the document.</p>
     *
     * @param doc         {@link XWPFDocument doc}
     * @param sourceTable {@link XWPFTable sourceTable}
     * @return {@link XWPFTable}
     */
    public static XWPFTable copyTable(XWPFDocument doc, XWPFTable sourceTable, boolean isTail) {
        if (doc == null || sourceTable == null) {
            throw new RuntimeException("The parameters passed in cannot be empty!");
        }
        // doc.getPosOfTable：What is obtained is the position of the table in the body
        // int tableIndex = doc.getPosOfTable(sourceTable);
        int tableIndex = doc.getTables().indexOf(sourceTable) + 1;
        if (isTail) {
            tableIndex = doc.getTables().size();
        }
        CTTbl newTbl = doc.getDocument().getBody().insertNewTbl(tableIndex);
        newTbl.set(sourceTable.getCTTbl());
        XWPFTable table = new XWPFTable(newTbl, doc);
        doc.insertTable(tableIndex, table);
        return table;
    }

    /**
     * Copy the content of the current line to the next line. If the next line is a newly added line,
     * then directly copy the entire XML of the current line to the next line. Otherwise, just copy
     * the content to the next line.
     * <p><b>Tips</b> </p>
     * <p>Cross table copying of the same file has not been tested, please use with caution</p>
     * <p>Cross table replication of different files has not been tested, please use with caution</p>
     *
     * @param currentLine      current line
     * @param nextLine         next line
     * @param templateRowIndex next line row index
     */
    public static XWPFTableRow copyLineContent(XWPFTableRow currentLine, XWPFTableRow nextLine, int templateRowIndex) {
        if (currentLine == null || nextLine == null) {
            return nextLine;
        }
        XWPFTable table = nextLine.getTable();
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
     * @param source         {@link XWPFTableCell source}
     * @param target         {@link XWPFTableCell target}
     * @param isIncludeStyle true: include style, false: not include style
     */
    public static void copyCellContent(@NotNull XWPFTableCell source, @NotNull XWPFTableCell target, boolean isIncludeStyle) {
        if (source == null || target == null) {
            return;
        }
        List<XWPFParagraph> paragraphs = source.getParagraphs();
        if (CollectionUtils.isEmpty(paragraphs)) {
            cleanCellContent(target);
            return;
        }
        CTPPr targetCtPPr = null;
        XWPFParagraph firstParagraph = null;
        if (org.apache.commons.collections4.CollectionUtils.isNotEmpty(target.getParagraphs())) {
            firstParagraph = target.getParagraphs().get(0);
            targetCtPPr = firstParagraph.getCTP().getPPr();
        }
        for (XWPFParagraph paragraph : source.getParagraphs()) {
            XWPFParagraph newParagraph = target.addParagraph();
            WordTableUtils.copyParagraph(paragraph, newParagraph, isIncludeStyle);
            if (targetCtPPr != null) {
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
                target.getCTTc().setTcPr(sourceTcPr);
            }
        }
        if (firstParagraph != null) {
            WordTableUtils.removeParagraphOfCell(target, firstParagraph);
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
        removeAllParagraphsOfCell(target);
        for (XWPFParagraph paragraph : source.getParagraphs()) {
            XWPFParagraph newParagraph = target.addParagraph();
            copyParagraph(paragraph, newParagraph, true);
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

    // copy paragraph content withouw style
    public static void copyParagraphContent(XWPFParagraph source, XWPFParagraph target) {
        copyParagraph(source, target, false);
    }

    /**
     * <p>Copy paragraph. The <b>same</b> document can be copied into any paragraph, and paragraphs
     * can also be copied <b>across</b> documents</p>
     * <p><b>Picture</b> paragraphs can also be copied!</p>
     *
     * @param source         {@link XWPFParagraph source}
     * @param target         {@link XWPFParagraph target}
     * @param isIncludeStyle true: include style, false: not include style
     * @see org.apache.poi.common.usermodel.PictureType
     */
    public static void copyParagraph(XWPFParagraph source, XWPFParagraph target, boolean isIncludeStyle) {
        if (target == null || source == null) {
            return;
        }
        cleanParagraphContent(target);
        if (CollectionUtils.isEmpty(source.getRuns())) {
            return;
        }
        XWPFDocument destDoc = target.getDocument();
        for (XWPFRun run : source.getRuns()) {
            XWPFRun newRun = target.createRun();
            // picture deal
            int id = 1;
            for (XWPFPicture picture : run.getEmbeddedPictures()) {
                try {
                    XWPFPictureData picData = picture.getPictureData();
                    byte[] pictureBytes = picData.getData();
                    int pictureFormat = picData.getPictureType();
                    CTPicture ctPicture = picture.getCTPicture();
                    // Adds a picture to the document.
                    String blipId = destDoc.addPictureData(pictureBytes, pictureFormat);
                    XWPFPicture newPicture = newRun.addPicture(new ByteArrayInputStream(pictureBytes), pictureFormat,
                        picData.getFileName(), Units.toEMU(picture.getWidth()), Units.toEMU(picture.getDepth()));
                    CTPicture newCTPicture = newPicture.getCTPicture();
                    newCTPicture.set(ctPicture);
                    // Connect image data to the a:blip element
                    newCTPicture.getBlipFill().getBlip().setEmbed(blipId);
                } catch (InvalidFormatException | IOException ignore) {
                }
            }
            copyRun(run, newRun, isIncludeStyle);
        }
        if (isIncludeStyle) {
            target.getCTP().setPPr(source.getCTPPr());
        }
    }

    public static void copyRun(XWPFRun source, XWPFRun target, boolean isIncludeStyle) {
        if (target == null || source == null) {
            return;
        }
        IRunBody parent = target.getParent();
        XWPFRun newRun;
        if (parent instanceof XWPFParagraph) {
            XWPFParagraph targetParent = (XWPFParagraph) parent;
            newRun = targetParent.createRun();
        } else if (parent instanceof XWPFStructuredDocumentTagContent) {
            XWPFStructuredDocumentTagContent targetParent = (XWPFStructuredDocumentTagContent) parent;
            newRun = targetParent.createRun();
        } else {
            logger.warn("XWPFRun's parent {} does not currently support processing", parent);
            return;
        }
        if (isIncludeStyle) {
            CTR sourceCTR = source.getCTR();
            if (sourceCTR.isSetRPr()) {
                CTRPr sourceCTRPr = sourceCTR.getRPr();
                CTRPr newCTRPr = CTRPr.Factory.newInstance();
                copyCTRPr(sourceCTRPr, newCTRPr);
                newRun.getCTR().setRPr(newCTRPr);
            }
        }
        newRun.setText(source.text());
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
        removeAllParagraphsOfCell(cell);
    }

    /**
     * Clear the content of XWPFParagraph
     *
     * @param paragraph {@link XWPFParagraph paragraph}
     */
    public static void cleanParagraphContent(XWPFParagraph paragraph) {
        if (paragraph == null) {
            return;
        }
        for (int i = paragraph.getRuns().size() - 1; i >= 0; i--) {
            paragraph.removeRun(i);
        }
    }

    public static void removeTable(XWPFDocument document, XWPFTable table) {
        if (table == null || document == null) {
            return;
        }
        int posOfTable = document.getPosOfTable(table);
        if (posOfTable < 0) {
            return;
        }
        document.removeBodyElement(document.getPosOfTable(table));
    }

    /**
     * Remove paragraph self
     *
     * @param paragraph {@link XWPFParagraph paragraph}
     */
    public static void removeParagraph(XWPFParagraph paragraph) {
        if (paragraph == null) {
            return;
        }
        if (!paragraph.getRuns().isEmpty()) {
            return;
        }
        IBody body = paragraph.getBody();
        if (body instanceof XWPFDocument) {
            XWPFDocument document = (XWPFDocument) body;
            int posOfTable = document.getPosOfParagraph(paragraph);
            document.removeBodyElement(posOfTable);
        } else if (body instanceof XWPFTableCell) {
            XWPFTableCell cell = (XWPFTableCell) body;
            cell.removeParagraph(cell.getParagraphs().indexOf(paragraph));
        } else if (body instanceof XWPFStructuredDocumentTagContent) {
            XWPFStructuredDocumentTagContent parent = (XWPFStructuredDocumentTagContent) body;
            parent.removeParagraph(paragraph);
        } else if (body instanceof XWPFComment) {
            XWPFComment parent = (XWPFComment) body;
            parent.removeParagraph(paragraph);
        } else if (body instanceof XWPFTextboxContent) {
            XWPFTextboxContent parent = (XWPFTextboxContent) body;
            parent.removeParagraph(paragraph);
        } else if (body instanceof XWPFHeaderFooter) {
            XWPFHeader parent = (XWPFHeader) body;
            parent.removeParagraph(paragraph);
        }
    }

    public static void removeBlankParagraph(XWPFDocument document) {
        // 
    }

    public static void removeAllParagraphsOfCell(XWPFTableCell cell) {
        if (cell == null) {
            return;
        }
        while (!cell.getParagraphs().isEmpty()) {
            cell.removeParagraph(0);
        }
    }

    public static void removeParagraphOfCell(XWPFTableCell cell, XWPFParagraph paragraph) {
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
     * Get the sum of all GridSpan in the longest row
     *
     * @param table {@link XWPFTable table}
     * @return int
     */
    public static int getTableMaxLineAllGridSpan(XWPFTable table) {
        if (table == null) {
            return 0;
        }

        int all = 0;

        // 遍历每一行
        for (XWPFTableRow row : table.getRows()) {
            all = 0;
            // 遍历每一列
            for (XWPFTableCell cell : row.getTableCells()) {
                CTTc ctTc = cell.getCTTc();
                CTTcPr tcPr = ctTc.isSetTcPr() ? ctTc.getTcPr() : ctTc.addNewTcPr();

                // 检查是否设置了 gridSpan
                if (tcPr.isSetGridSpan()) {
                    CTDecimalNumber gridSpan = tcPr.getGridSpan();
                    int currentGridSpan = gridSpan.getVal().intValue();
                    all += currentGridSpan;
                } else {
                    all++;
                }
            }
        }
        return all;
    }

    /**
     * obtain the count of vertically merged rows (Issue: If the columns are misaligned, the handling method has problems)
     *
     * @param table    XWPFTable
     * @param startRow start row index
     * @param colIndex col index
     * @return span col number，0 indicates no cross row
     */
    public static int findVerticalMergedRows(XWPFTable table, int startRow, int colIndex) {
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

    public static int findRowHeight(XWPFTableRow row) {
        return row.getHeight();
    }

    public static String findCellWidth(XWPFTableCell cell) {
        CTTcPr tcPr = cell.getCTTc().getTcPr();
        if (tcPr != null && tcPr.isSetTcW()) {
            return tcPr.getTcW().getW().toString();
        }
        return "0";
    }

    /**
     * Retrieve the spanned row data, where restart=2 indicates the start of a span.
     * continue=1 signifies the continuation of the spanned data, and the spanning ends when there is no more span information.
     *
     * @param cell {@link XWPFTableCell cell}
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

    public static void setTableWidthA4(XWPFTable table) {
        setTableWidth(table, Page.A4_NORMAL);
    }

    public static void setTableWidth(XWPFTable table, Page page) {
        CTTbl ctTbl = table.getCTTbl();
        CTTblPr tblPr = (ctTbl.getTblPr() != null) ? ctTbl.getTblPr() : ctTbl.addNewTblPr();
        CTTblWidth tblWidth = tblPr.addNewTblW();
        // tblWidth.setW(BigInteger.valueOf(5000));  // set 50%（5000 表示 50.00%）
        // tblWidth.setW(BigInteger.valueOf(10000));  // set 50%（5000 表示 100.00%）
        tblWidth.setW(page.contentWidth()); // Word unit is wips
        tblWidth.setType(STTblWidth.DXA);
    }

    @SuppressWarnings("unchecked")
    public static void setTableRow(XWPFTable table, XWPFTableRow row, int pos) {
        List<XWPFTableRow> rows = (List<XWPFTableRow>) ReflectionUtils.getValue("tableRows", table);
        rows.set(pos, row);
        table.getCTTbl().setTrArray(pos, row.getCtRow());
    }


    // The row itself does not have a direct width attribute
    @Deprecated
    public static void setTableRowWidth(XWPFTableRow row, int width) {
        throw new RuntimeException("The row itself does not have a direct width attribute");
    }

    /**
     * <p>Set row height</p>
     * Twips explanation:
     * Twip is a unit used in Microsoft Word, where 1 Twip equals 1/20 pound. Therefore,
     * 500 Twips is equivalent to 25 pounds (500 ÷ 20=25).
     *
     * @param row         {@link XWPFTableRow row}
     * @param heightTwips height in twips
     * @param type        {@link STHeightRule.Enum type} , default is {@link STHeightRule.Enum EXACT}
     */
    public static void setTableRowHeight(XWPFTableRow row, int heightTwips, STHeightRule.Enum type) {
        CTRow ctRow = row.getCtRow();
        CTTrPr ctTrPr = ctRow.isSetTrPr() ? ctRow.getTrPr() : ctRow.addNewTrPr();
        CTHeight height = ctTrPr.addNewTrHeight();
        height.setVal(BigInteger.valueOf(heightTwips));
        if (type == null) {
            type = STHeightRule.EXACT;
        }
        // Set the row height rule to a fixed row height
        height.setHRule(type);  // EXCT: indicates a fixed height
    }

    public static void setTableCellWidth(XWPFTableCell cell, String width) {
        CTTcPr tcPr = cell.getCTTc().isSetTcPr() ? cell.getCTTc().getTcPr() : cell.getCTTc().addNewTcPr();
        CTTblWidth tblWidth = tcPr.isSetTcW() ? tcPr.getTcW() : tcPr.addNewTcW();
        tblWidth.setType(STTblWidth.DXA);
        tblWidth.setW(new BigInteger(width));
    }

    public static void setCellWidth(XWPFTableCell cell, int width) {
        CTTblWidth tblWidth = CTTblWidth.Factory.newInstance();
        tblWidth.setType(STTblWidth.DXA);
        tblWidth.setW(BigInteger.valueOf(width));
        CTTcPr tcPr = cell.getCTTc().isSetTcPr() ? cell.getCTTc().getTcPr() : cell.getCTTc().addNewTcPr();
        tcPr.setTcW(tblWidth);
    }

    public static void setDiagonalBorder(XWPFTableCell cell) {
        setDiagonalBorder(cell, STBorder.SINGLE);
    }

    /**
     * <p>Set diagonal border (slash)</p>
     * <p>Expand :</p>
     * <p>1. Tr2bl: The diagonal border from the top right corner (tr) to the bottom left corner (bl) of the cell.</p>
     * <p>2. Tl2br: The diagonal border from the top left corner (tl) to the bottom right corner (br) of the cell.</p>
     *
     * @param cell     {@link XWPFTableCell cell}
     * @param stBorder {@link STBorder.Enum stBorder} line style
     */
    public static void setDiagonalBorder(XWPFTableCell cell, STBorder.Enum stBorder) {
        if (cell == null) {
            return;
        }
        CTTc ctTc = cell.getCTTc();
        CTTcPr ctTcPr = ctTc.isSetTcPr() ? ctTc.getTcPr() : ctTc.addNewTcPr();
        CTTcBorders ctTcBorders = ctTcPr.isSetTcBorders() ? ctTcPr.getTcBorders() : ctTcPr.addNewTcBorders();
        if (!ctTcBorders.isSetTr2Bl()) {
            ctTcBorders.addNewTr2Bl().setVal(stBorder);
        }
    }

    public static void setPageBreak(XWPFDocument document) {
        XWPFParagraph pageBreakPara = document.createParagraph();
        pageBreakPara.setPageBreak(true);
        // XWPFRun pageBreakRun = pageBreakPara.createRun();
        // pageBreakRun.addBreak(BreakType.PAGE);
    }

    /**
     * <p>XWPFParagraph.setPageBreak(true) sets a page break at the paragraph level. It will move the entire paragraph
     * content to a new page</p>
     * <p>XWPFRun.addBreak(BreakType.PAGE) inserts a page break in the text run (XWPFRun), which causes the page break
     * to be inserted from the current text position and the subsequent content is moved to a new page</p>
     *
     * @param document {@link XWPFDocument document}
     * @param body     {@link  IBodyElement body}
     */
    public static void setPageBreak(XWPFDocument document, IBodyElement body) {
        if (document == null || body == null) {
            return;
        }
        XmlObject xmlObject = null;
        if (body.getElementType() == BodyElementType.PARAGRAPH) {
            XWPFParagraph paragraph = (XWPFParagraph) body;
            xmlObject = paragraph.getCTP();
        } else if (body.getElementType() == BodyElementType.TABLE) {
            XWPFTable table = (XWPFTable) body;
            xmlObject = table.getCTTbl();
        }
        if (xmlObject == null) {
            return;
        }
        XmlCursor xmlCursor = xmlObject.newCursor();
        xmlCursor.toNextSibling();
        XWPFParagraph pageBreakPara = document.insertNewParagraph(xmlCursor);
        pageBreakPara.setPageBreak(true);
//        XWPFRun pageBreakRun = pageBreakPara.createRun();
//        pageBreakRun.addBreak(BreakType.PAGE);
    }

    public static void mergeCellsHorizontalFullLine(XWPFTableRow tableRow) {
        if (tableRow == null) {
            return;
        }
        mergeCellsHorizontal(tableRow.getTable(), 0, tableRow, 0, tableRow.getTableCells().size() - 1);
    }

    public static void mergeCellsHorizontalFullLine(XWPFTable table, int row) {
        if (table == null) {
            return;
        }
        int size = table.getRows().size();
        if (size <= row) {
            throw new RuntimeException("row index out of bounds");
        }
        XWPFTableRow tableRow = table.getRow(row);
        size = tableRow.getTableCells().size();
        mergeCellsHorizontal(table, row, tableRow, 0, size - 1);
    }

    /**
     * <p>Merge cells with specified start and end position <b>indexes</b> in a row.</p>
     * <p>The priority of the <b>tableRow</b> incoming row object is higher than the specified <b>row</b> index row</p>
     *
     * @param table    {@link XWPFTable table} table object
     * @param row      row index
     * @param tableRow {@link XWPFTableRow tableRow} row object
     * @param fromCol  from column index
     * @param toCol    to column index
     */
    public static void mergeCellsHorizontal(XWPFTable table, int row, XWPFTableRow tableRow, int fromCol, int toCol) {
        if (table == null) {
            return;
        }

        List<XWPFTableRow> xwpfRow = table.getRows();
        int rowCount = xwpfRow.size();

        if (row < 0 || row >= rowCount) {
            throw new IllegalArgumentException(row + " index out of bounds");
        }
        if (tableRow == null) {
            tableRow = table.getRow(row);
        }
        if (tableRow == null) {
            throw new IllegalArgumentException(row + " not existed");
        }

        List<XWPFTableCell> cells = tableRow.getTableCells();
        int cellCount = cells.size();

        if (toCol < fromCol || fromCol < 0 || toCol >= cellCount) {
            throw new IllegalArgumentException("col index out of bounds");
        }

        int tableMaxLineAllGridSpan = getTableMaxLineAllGridSpan(table);
        XWPFTableCell startCell = tableRow.getCell(fromCol);
        CTTc startCTTc = startCell.getCTTc();
        CTTcPr startTcPr = startCTTc.isSetTcPr() ? startCTTc.getTcPr() : startCTTc.addNewTcPr();
        if (startTcPr.isSetVMerge()) {
            startTcPr.unsetVMerge();
        }

        CTHMerge hMerge = startTcPr.isSetHMerge() ? startTcPr.getHMerge() : startTcPr.addNewHMerge();
        hMerge.setVal(STMerge.RESTART);
        // set gridSpan element
        CTDecimalNumber gridSpan = startTcPr.isSetGridSpan() ? startTcPr.getGridSpan() : startTcPr.addNewGridSpan();
        gridSpan.setVal(BigInteger.valueOf(tableMaxLineAllGridSpan));

        for (int colIndex = fromCol + 1; colIndex <= toCol; colIndex++) {
            XWPFTableCell cell = tableRow.getCell(colIndex);
            CTTc ctTc = cell.getCTTc();
            CTTcPr tcPr = ctTc.isSetTcPr() ? ctTc.getTcPr() : ctTc.addNewTcPr();
            if (tcPr.isSetVMerge()) {
                tcPr.unsetVMerge();
            }
            CTHMerge continueHMerge = tcPr.isSetHMerge() ? tcPr.getHMerge() : tcPr.addNewHMerge();
            continueHMerge.setVal(STMerge.CONTINUE);
            cleanCellContent(cell);
        }
    }

    public static void unmergeCells(XWPFTableRow row, int startCellIndex, boolean isAddSpan) {
        if (row == null) {
            return;
        }
        if (startCellIndex < 0 || startCellIndex >= row.getTableCells().size()) {
            throw new IndexOutOfBoundsException("Invalid startCellIndex: " + startCellIndex);
        }
        XWPFTableCell cell = row.getCell(startCellIndex);
        if (cell == null) {
            return;
        }
        CTTcPr tcPr = cell.getCTTc().getTcPr();
        if (tcPr == null) {
            return;
        }
        if (tcPr.isSetHMerge()) {
            tcPr.unsetHMerge();
        }
        if (tcPr.isSetVMerge()) {
            CTVMerge vMerge = tcPr.getVMerge();
            if (vMerge != null) {
                tcPr.unsetVMerge();
            }
        }
        if (tcPr.isSetGridSpan()) {
            CTDecimalNumber gridSpan = tcPr.getGridSpan();
            if (gridSpan != null && gridSpan.getVal() != null) {
                int gridSpanValue = gridSpan.getVal().intValue();
                if (gridSpanValue > 1) {
                    tcPr.unsetGridSpan();
                    if (isAddSpan) {
                        for (int i = 1; i < gridSpanValue; i++) {
                            addCellAfter(row, startCellIndex);
                        }
                    }
                }
            }
        }
    }

    public static void unVMergeCells(XWPFTableRow row, int startCellIndex) {
        if (row == null) {
            return;
        }
        XWPFTableCell cell = row.getCell(startCellIndex);
        if (cell == null) {
            return;
        }
        CTTcPr tcPr = cell.getCTTc().getTcPr();
        if (tcPr == null) {
            return;
        }
        if (tcPr.isSetVMerge()) {
            CTVMerge vMerge = tcPr.getVMerge();
            if (vMerge != null) {
                tcPr.unsetVMerge();
            }
        }
    }

    /**
     * Merge a column into a single cell by specifying the start and end rows.
     *
     * @param table   {@link XWPFTable table} table object
     * @param col     column index
     * @param fromRow from column index
     * @param toRow   to column index
     */
    public static void mergeCellsVertically(XWPFTable table, int col, int fromRow, int toRow) {
        if (table == null || fromRow < 0 || toRow >= table.getRows().size() || col < 0) {
            return;
        }
        for (int rowIndex = fromRow; rowIndex <= toRow; rowIndex++) {
            XWPFTableRow row = table.getRow(rowIndex);
            if (row == null) {
                continue;
            }
            XWPFTableCell cell = row.getCell(col);
            if (cell == null) {
                continue;
            }
            CTTc ctTc = cell.getCTTc();
            CTTcPr ctTcPr = ctTc.isSetTcPr() ? ctTc.getTcPr() : ctTc.addNewTcPr();
            CTVMerge ctvMerge = ctTcPr.isSetVMerge() ? ctTcPr.getVMerge() : ctTcPr.addNewVMerge();
            ctvMerge.setVal(rowIndex == fromRow ? STMerge.RESTART : STMerge.CONTINUE);
        }
    }

    /**
     * <p>Merge multiple rows. Merge each row into a column, and then merge multiple rows. Finally, retain the original
     * text content of the first row and first column </p>
     * <p>This merge does not actually merge the parallel numbers into one row and one column,
     * but is displayed as one row</p>
     *
     * @param table   {@link XWPFTable table}
     * @param fromRow from row index
     * @param toRow   to row index
     */
    public static void mergeMutipleLine(XWPFTable table, int fromRow, int toRow) {
        if (table == null) {
            return;
        }
        List<XWPFTableRow> rows = table.getRows();
        if (rows.size() <= toRow || fromRow < 0 || fromRow > toRow) {
            throw new RuntimeException(String.format("The input row index(%d,%d) is incorrect", fromRow, toRow));
        }
        for (int i = fromRow; i <= toRow; i++) {
            XWPFTableRow row = rows.get(i);
            mergeCellsHorizontal(table, i, row, 0, row.getTableCells().size() - 1);
        }
        mergeCellsVertically(table, 0, fromRow, toRow);
    }

    /**
     * Add a new cell after a specified row cell
     *
     * @param row       a specified row
     * @param cellIndex Add a new cell after this cell
     */
    public static void addCellAfter(XWPFTableRow row, int cellIndex) {
        if (row == null) {
            return;
        }
        XWPFTable table = row.getTable();
        int rowSize = table.getRows().size();
        int colSize = row.getTableCells().size();
        int curerntRowIndex = findRowIndex(row);

        XWPFTableCell newCell = row.addNewTableCell();

        for (int i = colSize; i > cellIndex + 1; i--) {
            XWPFTableCell tempCell = row.getCell(i - 1);
            row.getCtRow().setTcArray(i, row.getCtRow().getTcArray(i - 1));
            row.getCtRow().setTcArray(i - 1, tempCell.getCTTc());
        }
    }
}
