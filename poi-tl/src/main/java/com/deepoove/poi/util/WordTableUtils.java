package com.deepoove.poi.util;

import com.deepoove.poi.xwpf.Page;
import com.deepoove.poi.xwpf.XWPFStructuredDocumentTag;
import com.deepoove.poi.xwpf.XWPFStructuredDocumentTagContent;
import com.deepoove.poi.xwpf.XWPFTextboxContent;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.wp.usermodel.Paragraph;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlObject;
import org.openxmlformats.schemas.drawingml.x2006.picture.CTPicture;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.util.List;

/**
 * copy → clean → remove → find → set → merge
 * TODO updat doc property id， DrawingSupport.updateDocPrId(table);
 */
@SuppressWarnings("unused")
public class WordTableUtils {

    private static final Logger logger = LoggerFactory.getLogger(WordTableUtils.class);

    public static XWPFTable copyTable(XWPFDocument doc, XWPFTable sourceTable) {
        return copyTable(doc, sourceTable, null, false);
    }

    public static XWPFTable copyTable(XWPFDocument doc, XWPFTable sourceTable, boolean isTail) {
        return copyTable(doc, sourceTable, null, isTail);
    }

    public static XWPFTable copyTable(XWPFDocument doc, XWPFTable sourceTable, XmlCursor xmlCursor) {
        return copyTable(doc, sourceTable, xmlCursor, false);
    }

    /**
     * <p>Copy the specified table after the specified {@link XmlCursor xmlCursor}. If <b>{@link XmlCursor xmlCursor}</b> parameter is <b>empty</b>,
     * it will be added by default after the position of the <b>current</b> tbale.</p>
     * <p>If isTail is True, the new table will be placed at the end of the document.</p>
     * <p>If the {@link XmlCursor xmlCursor} parameter is created externally, remember to release the resource yourself</p>
     *
     * @param doc         {@link XWPFDocument doc}
     * @param sourceTable {@link XWPFTable sourceTable}
     * @param xmlCursor   {@link XmlCursor xmlCursor} Add a table after the specified XMLCursor
     * @param isTail      boolean, whether to place the new table at the end of the document
     * @return {@link XWPFTable}
     */
    @SuppressWarnings("unchecked")
    public static XWPFTable copyTable(XWPFDocument doc, XWPFTable sourceTable, XmlCursor xmlCursor, boolean isTail) {
        if (doc == null || sourceTable == null) {
            throw new RuntimeException("The parameters passed in cannot be empty!");
        }
        // doc.getPosOfTable：What is obtained is the position of the table in the body
        // int tableIndex = doc.getPosOfTable(sourceTable);
        if (xmlCursor == null) {
            xmlCursor = sourceTable.getCTTbl().newCursor();
        }
        xmlCursor.toNextSibling();
        if (isTail) {
            while (xmlCursor.toNextSibling()) ;
        }
        XWPFTable table = doc.insertNewTbl(xmlCursor);
        table.removeRow(0);
        CTTbl ctTbl = table.getCTTbl();
        ctTbl.set(sourceTable.getCTTbl());
        CTRow[] trArray = ctTbl.getTrArray();
        List<XWPFTableRow> tableRows = (List<XWPFTableRow>) ReflectionUtils.getValue("tableRows", table);
        for (int i = 0; i < trArray.length; i++) {
            XWPFTableRow row = new XWPFTableRow(trArray[i], table);
            tableRows.add(i, row);
        }
        return table;
    }

    /**
     * <p>Copy the table style of the source table to the target table.</p>
     *
     * @param sourceTable {@link  XWPFTable sourceTable}
     * @param targetTable {@link  XWPFTable targetTable}
     */
    public static void copyTableTblPr(XWPFTable sourceTable, XWPFTable targetTable) {
        if (sourceTable == null || targetTable == null) {
            return;
        }
        CTTblPr sourceTblPr = sourceTable.getCTTbl().getTblPr();
        CTTblPr targetTblPr = targetTable.getCTTbl().getTblPr();
        targetTblPr.set(sourceTblPr);
    }

    /**
     * <p>TODO Copy the content of the current line to the next line.</p>
     *
     * @param currentLine    {@link XWPFTableRow currentLine}
     * @param nextLine       {@link  XWPFTableRow nextLine}
     * @param isIncludeStyle boolean, whether to keep the style of the target cell
     */
    public static void copyLine(XWPFTableRow currentLine, XWPFTableRow nextLine, boolean isIncludeStyle) {
        // if in the same document
        XWPFTable sourceTable = currentLine.getTable();
        XWPFTable targetTable = currentLine.getTable();

        if (sourceTable == targetTable) {
        }
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
        if (CollectionUtils.isEmpty(nextLine.getTableCells())) {
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
     * Copy cell data to another cell, and if the source cell does not have data, clear the target cell data.
     * course,
     *
     * @param source         {@link XWPFTableCell source}
     * @param target         {@link  XWPFTableCell target}
     * @param isIncludeStyle boolean, whether to keep the style of the target cell
     */
    public static void copyCell(XWPFTableCell source, XWPFTableCell target, boolean isIncludeStyle) {
        if (source == null || target == null) {
            return;
        }
        List<XWPFParagraph> paragraphs = source.getParagraphs();
        cleanCellContent(target);
        if (CollectionUtils.isEmpty(paragraphs)) {
            return;
        }
        for (XWPFParagraph paragraph : paragraphs) {
            XWPFParagraph newParagraph = target.addParagraph();
            WordTableUtils.copyParagraph(paragraph, newParagraph, isIncludeStyle);
        }
        if (isIncludeStyle) {
            CTTcPr sourceTcPr = source.getCTTc().getTcPr();
            if (sourceTcPr != null) {
                target.getCTTc().setTcPr(sourceTcPr);
            }
        }
    }

    /**
     * Copy cell data to another cell, but keep the style of the target cell. The isIncludeStyle field is used to
     * control whether to overwrite the source cell style
     *
     * @param source         {@link XWPFTableCell source}
     * @param target         {@link XWPFTableCell target}
     * @param isIncludeStyle true: include style, false: not include style
     */
    public static void copyCellContent(XWPFTableCell source, XWPFTableCell target, boolean isIncludeStyle) {
        if (source == null || target == null) {
            return;
        }
        CTPPr targetCtPPr = null;
        XWPFParagraph firstParagraph = null;
        if (CollectionUtils.isNotEmpty(target.getParagraphs())) {
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
        XWPFDocument sourceDoc = source.getDocument();
        boolean isSameDoc = sourceDoc == destDoc;
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
                    String blipId = isSameDoc ? ctPicture.getBlipFill().getBlip().getEmbed() : destDoc.addPictureData(pictureBytes, pictureFormat);
                    XWPFPicture newPicture = newRun.addPicture(new ByteArrayInputStream(pictureBytes), pictureFormat, picData.getFileName(), Units.toEMU(picture.getWidth()), Units.toEMU(picture.getDepth()));
                    CTPicture newCTPicture = newPicture.getCTPicture();
                    if (isIncludeStyle) {
                        newCTPicture.set(ctPicture);
                    }
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
        if (isIncludeStyle) {
            CTR sourceCTR = source.getCTR();
            if (sourceCTR.isSetRPr()) {
                CTRPr sourceCTRPr = sourceCTR.getRPr();
                CTRPr newCTRPr = CTRPr.Factory.newInstance();
                copyCTRPr(sourceCTRPr, newCTRPr);
                target.getCTR().setRPr(newCTRPr);
            }
        }
        target.setText(source.text());
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
            List<XWPFParagraph> paragraphs = cell.getParagraphs();
            for (int i = paragraphs.size() - 1; i >= 0; i--) {
                cell.removeParagraph(i);
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

    public static void removeRow(XWPFTable table, int rowIndex) {
        XWPFTableRow row = table.getRow(rowIndex);
        if (row == null) {
            logger.warn("rowIndex is out of range");
            return;
        }
        removeRow(row);
    }

    public static void removeRow(XWPFTableRow row) {
        if (row == null) {
            return;
        }
        XWPFTable table = row.getTable();
        int rowIndex = table.getRows().indexOf(row);
        table.removeRow(rowIndex);
    }

    public static void removeLastRow(XWPFTable table) {
        removeRow(table, table.getRows().size() - 1);
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

    /**
     * Remove the last blank paragraph from the document
     *
     * @param xwpfDocument {@link XWPFDocument xwpfDocument}
     */
    public static void removeLastBlankParagraph(XWPFDocument xwpfDocument) {
        if (xwpfDocument == null) {
            return;
        }
        List<IBodyElement> bodyElements = xwpfDocument.getBodyElements();
        // The last few lines are empty, delete elements to avoid creating a new page
        for (int i = bodyElements.size() - 1; i >= 0; i--) {
            IBodyElement iBodyElement = bodyElements.get(i);
            if (iBodyElement instanceof XWPFParagraph) {
                XWPFParagraph paragraph = (XWPFParagraph) iBodyElement;
                if (CollectionUtils.isEmpty(paragraph.getRuns())) {
                    xwpfDocument.removeBodyElement(i);
                } else {
                    List<XWPFRun> runs = paragraph.getRuns();
                    boolean isEmpty = true;
                    for (XWPFRun run : runs) {
                        if (StringUtils.isNotBlank(run.text())) {
                            isEmpty = false;
                            break;
                        }
                    }
                    if (isEmpty) {
                        xwpfDocument.removeBodyElement(i);
                    }
                }
            } else {
                break;
            }
        }
    }

    public static void removeRun(XWPFParagraph paragraph, XWPFRun run) {
        if (!CollectionUtils.isEmpty(paragraph.getRuns())) {
            paragraph.removeRun(paragraph.getRuns().indexOf(run));
        }
    }

    public static void removeAllRun(XWPFParagraph paragraph) {
        if (paragraph != null && CollectionUtils.isNotEmpty(paragraph.getRuns())) {
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
    public static int findTableMaxLineAllGridSpan(XWPFTable table) {
        if (table == null) {
            return 0;
        }
        CTTblGrid tblGrid = table.getCTTbl().getTblGrid();
        int all = 0;
        if (tblGrid == null) {
            for (XWPFTableRow row : table.getRows()) {
                int temp = 0;
                for (XWPFTableCell cell : row.getTableCells()) {
                    CTTc ctTc = cell.getCTTc();
                    CTTcPr tcPr = ctTc.isSetTcPr() ? ctTc.getTcPr() : ctTc.addNewTcPr();
                    if (tcPr.isSetGridSpan()) {
                        CTDecimalNumber gridSpan = tcPr.getGridSpan();
                        int currentGridSpan = gridSpan.getVal().intValue();
                        temp += currentGridSpan;
                    } else {
                        temp++;
                    }
                }
                if (temp > all) {
                    all = temp;
                }
            }
        } else {
            all = tblGrid.getGridColList().size();
        }
        return all;
    }

    /**
     * Get the table margin
     *
     * @param table {XWPFTable table}
     * @param flag  flag equals 1: top, 2: bottom
     * @return long
     */
    public static int findTableMargin(XWPFTable table, int flag) {
        CTTblPr tblPr = table.getCTTbl().getTblPr();
        if (tblPr != null) {
            CTTblCellMar tblCellMar = tblPr.getTblCellMar();
            BigInteger bigInteger = new BigInteger("0");
            if (tblCellMar == null) {
                return 0;
            }
            if (flag == 1) {
                CTTblWidth topMar = tblCellMar.getTop();
                if (topMar != null) {
                    bigInteger = (BigInteger) topMar.getW();
                }
            } else {
                CTTblWidth bottomMar = tblCellMar.getBottom();
                if (bottomMar != null) {
                    bigInteger = (BigInteger) bottomMar.getW();
                }
            }
            return bigInteger.intValue();
        } else {
            return 0;
        }
    }

    /**
     * obtain the count of vertically merged rows (Issue: If the columns are misaligned, the handling method has problems)
     *
     * @param table    {@link XWPFTable table}
     * @param startRow start row index
     * @param colIndex col index
     * @return span col number，0 indicates no cross row
     */
    public static int findVerticalMergedRows(XWPFTable table, int startRow, int colIndex) {
        if (table == null) {
            return 1;
        }
        int i = startRow + 1;
        int size = table.getRows().size();
        for (; i < size; i++) {
            if (table.getRow(i).getCell(colIndex) == null) {
                break;
            }
            XWPFTableCell xwpfTableCell = table.getRow(i).getCell(colIndex);
            if (xwpfTableCell == null || xwpfTableCell.getCTTc() == null) {
                break;
            }
            CTTc ctTc = xwpfTableCell.getCTTc();
            if (!ctTc.isSetTcPr()) {
                break;
            }
            CTTcPr tcPr = ctTc.getTcPr();
            if (!tcPr.isSetVMerge()) {
                break;
            }
            CTVMerge vMerge = tcPr.getVMerge();
            if (vMerge == null || vMerge.getVal() == STMerge.RESTART) {
                break;
            }
        }
        return i - startRow;
    }

    public static int findVerticalMergedRows(XWPFTable table, XWPFTableCell cell) {
        if (cell == null || table == null) {
            return 1;
        }
        XWPFTableRow tableRow = cell.getTableRow();
        int rowIndex = findRowIndex(tableRow);
        int colIndex = tableRow.getTableCells().indexOf(cell);
        return findVerticalMergedRows(table, rowIndex, colIndex);
    }

    public static XWPFTableRow findLastLine(XWPFTable table) {
        if (table == null) {
            return null;
        }
        return table.getRow(table.getRows().size() - 1);
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
     * Get the <b>cross row</b> count of a cell
     *
     * @param cell {@link XWPFTableCell cell}
     * @return int. If cell is null, return 0.
     */
    public static int findCellVMergeNumber(XWPFTableCell cell) {
        if (cell == null) {
            return 0;
        }
        XWPFDocument xwpfDocument = cell.getXWPFDocument();
        XWPFTable table = cell.getTableRow().getTable();
        int rowIndex = WordTableUtils.findRowIndex(cell);
        return WordTableUtils.findVerticalMergedRows(table, cell);
    }

    /**
     * Retrieve the spanned row data, where restart=2 indicates the start of a span.
     * continue=1 signifies the continuation of the spanned data, and the spanning ends when there is no more span information.
     * <p>CTVMerge directly returns the corresponding 1 or 2 if there is a value, and returns 0 in other cases</p>
     *
     * @param cell {@link XWPFTableCell cell}
     * @return Integer | null则表示没有跨行
     */
    public static int findVMerge(XWPFTableCell cell) {
        // Get cell properties
        if (cell == null) {
            return 0;
        }
        CTTcPr tcPr = cell.getCTTc().getTcPr();
        if (tcPr != null) {
            // Get vertical merge properties
            CTVMerge vMerge = tcPr.getVMerge();
            if (vMerge != null && vMerge.isSetVal()) {
                return vMerge.getVal().intValue();
            }
        }
        return 0;
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

    public double findFontSize(XWPFRun run) {
        if (run == null) {
            return 0;
        }
        return run.getFontSizeAsDouble();
    }

    /**
     * <p>Set element positions in XWPFDocument</p>
     * <p>TODO Move error：XmlValueDisconnectedException，Do not use</p>
     *
     * @param document    {@link XWPFDocument doc}
     * @param bodyElement {@link IBodyElement bodyElement}
     * @param position    The position of elements in IBodyElement
     */
    @SuppressWarnings("unchecked")
    public static void setElementPostion(XWPFDocument document, IBodyElement bodyElement, int position) {
        if (document == null || bodyElement == null) {
            return;
        }
        List<IBodyElement> bodyElements = (List<IBodyElement>) ReflectionUtils.getValue("bodyElements", document);
        if (position < 0 || position > bodyElements.size()) {
            throw new RuntimeException("The position of the element is out of range");
        }
        int index = bodyElements.indexOf(bodyElement);
        if (index < 0 || index == position) {
            return;
        }
        // Move XML elements
        IBodyElement oldBodyElement = bodyElements.get(position);
        XmlCursor oldXmlCursor = getXmlCursor(oldBodyElement);
        XmlCursor xmlCursor = getXmlCursor(bodyElement);
        if (oldXmlCursor != null && xmlCursor != null) {
            if (position == bodyElements.size()) {
                oldXmlCursor.toEndToken();
            }
            xmlCursor.moveXml(oldXmlCursor);
            oldXmlCursor.close();
            xmlCursor.close();
        }


        CTDocument1 ctDocument = document.getDocument();
        int order = 0;
        for (IBodyElement element : bodyElements) {
            if (element.equals(bodyElement)) {
                break;
            } else {
                if (element.getElementType() == bodyElement.getElementType()) {
                    ++order;
                }
            }
        }
        CTBody body = ctDocument.getBody();
        switch (bodyElement.getElementType()) {
            case PARAGRAPH:
                XWPFParagraph paragraph = (XWPFParagraph) bodyElement;
                List<Paragraph> paragraphs = (List<Paragraph>) ReflectionUtils.getValue("paragraphs", document);
                int i = paragraphs.indexOf(paragraph);
                if (i != order) {
                    List<CTP> pList = body.getPList();
                    pList.set(order, paragraph.getCTP());
                    paragraphs.remove(paragraph);
                    paragraphs.add(order, paragraph);
                }
                break;
            case TABLE:
                XWPFTable table = (XWPFTable) bodyElement;
                List<XWPFTable> tables = (List<XWPFTable>) ReflectionUtils.getValue("tables", document);
                int i1 = tables.indexOf(table);
                if (i1 != order) {
                    List<CTTbl> tblList = body.getTblList();
                    tblList.set(order, table.getCTTbl());
                    tables.remove(table);
                    tables.add(order, table);
                }
                break;
        }
    }

    private static XmlCursor getXmlCursor(IBodyElement bodyElement) {
        if (bodyElement.getElementType() == BodyElementType.PARAGRAPH) {
            return ((XWPFParagraph) bodyElement).getCTP().newCursor();
        } else if (bodyElement.getElementType() == BodyElementType.TABLE) {
            return ((XWPFTable) bodyElement).getCTTbl().newCursor();
        }
        return null;
    }

    public static void setTablePosition(XWPFDocument doc, XWPFTable targetTable, int tableIndex) {
        doc.insertTable(tableIndex, targetTable);
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
     * Twip is a unit used in Microsoft Word, where 1 Twip equals 1/20 point. Therefore,
     * 500 Twips is equivalent to 25 points (500 ÷ 20=25).
     *
     * @param row         {@link XWPFTableRow row}
     * @param heightTwips height in twips
     * @param type        {@link STHeightRule.Enum type} , default is {@link STHeightRule.Enum EXACT}
     */
    public static void setTableRowHeight(XWPFTableRow row, long heightTwips, STHeightRule.Enum type) {
        CTRow ctRow = row.getCtRow();
        CTTrPr ctTrPr = ctRow.isSetTrPr() ? ctRow.getTrPr() : ctRow.addNewTrPr();
        CTHeight height = ctTrPr.sizeOfTrHeightArray() == 0 ? ctTrPr.addNewTrHeight() : ctTrPr.getTrHeightArray(0);
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

    /**
     * <p>Insert a page break at the <b>end</b> of the paragraph</p>
     *
     * @param document   {@link XWPFDocument document}
     * @param pageMethod Paging method, 1 is the entire paragraph paginated to the next page, not equal to 1 is
     *                   the current paragraph content paginated to the next line, suitable for table pagination
     */
    public static void setPageBreakInLast(XWPFDocument document, int pageMethod) {
        XWPFParagraph pageBreakPara = document.createParagraph();
        if (pageMethod == 1) {
            pageBreakPara.setPageBreak(true);
        } else {
            XWPFRun pageBreakRun = pageBreakPara.createRun();
            pageBreakRun.addBreak(BreakType.PAGE);
        }
    }

    /**
     * <p>XWPFParagraph.setPageBreak(true) sets a page break at the paragraph level. It will move the entire paragraph
     * content to a new page</p>
     * <p>XWPFRun.addBreak(BreakType.PAGE) inserts a page break in the text run (XWPFRun), which causes the page break
     * to be inserted from the current text position and the subsequent content is moved to a new page</p>
     * <p><b>Hint</b>: It will generate a <b> blank line</b></p>
     *
     * @param document   {@link XWPFDocument document}
     * @param body       {@link  IBodyElement body}
     * @param pageMethod Paging method, 1 is the entire paragraph paginated to the next page, not equal to 1 is
     *                   the current paragraph content paginated to the next line, suitable for table pagination
     */
    public static XWPFParagraph setPageBreak(XWPFDocument document, IBodyElement body, int pageMethod) {
        if (document == null || body == null) {
            return null;
        }
        XmlObject xmlObject = null;
        if (body.getElementType() == BodyElementType.PARAGRAPH) {
            XWPFParagraph paragraph = (XWPFParagraph) body;
            xmlObject = paragraph.getCTP();
        } else if (body.getElementType() == BodyElementType.TABLE) {
            XWPFTable table = (XWPFTable) body;
            xmlObject = table.getCTTbl();
        } else if (body.getElementType() == BodyElementType.CONTENTCONTROL) {
            XWPFStructuredDocumentTag sdt = (XWPFStructuredDocumentTag) body;
            xmlObject = sdt.getCtSdtBlock();
        }
        if (xmlObject == null) {
            return null;
        }
        XmlCursor xmlCursor = xmlObject.newCursor();
        xmlCursor.toNextSibling();
        XWPFParagraph pageBreakPara = document.insertNewParagraph(xmlCursor);
        if (pageMethod == 1) {
            pageBreakPara.setPageBreak(true);
        } else {
            XWPFRun pageBreakRun = pageBreakPara.createRun();
            pageBreakRun.addBreak(BreakType.PAGE);
        }
        return pageBreakPara;
    }

    /**
     * Add page breaks to existing paragraphs
     *
     * @param pageBreakPara {@link XWPFParagraph pageBreakPara}
     * @param pageMethod    Paging method, 1 is the entire paragraph paginated to the next page, not equal to 1 is the
     *                      current paragraph content paginated to the next line, suitable for table pagination
     */
    public static void setPageBreak(XWPFParagraph pageBreakPara, int pageMethod) {
        if (pageBreakPara == null) {
            return;
        }
        if (pageMethod == 1) {
            pageBreakPara.setPageBreak(true);
        } else {
            XWPFRun pageBreakRun = pageBreakPara.createRun();
            pageBreakRun.addBreak(BreakType.PAGE);
        }
    }

    public static void setMinHeightParagraph(XWPFDocument document) {
        XWPFParagraph paragraph = document.createParagraph();
        setMinHeightParagraph(paragraph);
    }

    public static void setMinHeightParagraph(XWPFParagraph paragraph) {
        CTP ctp = paragraph.getCTP();
        CTPPr pPr = ctp.isSetPPr() ? ctp.getPPr() : ctp.addNewPPr();
        CTSpacing spacing = pPr.isSetSpacing() ? pPr.getSpacing() : pPr.addNewSpacing();
        spacing.setBefore(BigInteger.valueOf(0));
        spacing.setAfter(BigInteger.valueOf(0));
        spacing.setLine(BigInteger.valueOf(1));
        spacing.setLineRule(STLineSpacingRule.EXACT);
    }

    /**
     * Set the bottom border of the table to the default left border style or the left border style of the first cell
     *
     * @param table {@link XWPFTable table}
     */
    public static void setBottomBorder(XWPFTable table, CTBorder border) {
        if (table == null) {
            return;
        }
        if (border == null) {
            CTTblPr tblPr = table.getCTTbl().getTblPr();
            if (tblPr != null) {
                if (tblPr.isSetTblBorders()) {
                    CTTblBorders ctTblBorders = tblPr.getTblBorders();
                    if (ctTblBorders.isSetLeft()) {
                        border = ctTblBorders.getLeft();
                    }
                }
            }
        }
        XWPFTableRow row = WordTableUtils.findLastLine(table);
        if (row == null || CollectionUtils.isEmpty(row.getTableCells())) {
            return;
        } else {
            if (border == null || border.getSz() == null) {
                XWPFTableCell cell = row.getCell(0);
                CTTc ctTc = cell.getCTTc();
                if (ctTc.isSetTcPr()) {
                    CTTcPr ctTcPr = ctTc.getTcPr();
                    if (ctTcPr.isSetTcBorders()) {
                        CTTcBorders tcBorders = ctTcPr.getTcBorders();
                        if (tcBorders.isSetLeft()) {
                            border = tcBorders.getLeft();
                        }
                    }
                }
            }
        }
        if (border == null) {
            return;
        }

        // Set each cell in the row
        for (XWPFTableCell cell : row.getTableCells()) {
            CTTcPr tcPr = cell.getCTTc().getTcPr();
            if (tcPr == null) {
                tcPr = cell.getCTTc().addNewTcPr();
            }
            CTTcBorders borders = tcPr.getTcBorders();
            if (borders == null) {
                borders = tcPr.addNewTcBorders();
            }
            CTBorder bottomBorder = borders.getBottom();
            if (bottomBorder == null) {
                bottomBorder = borders.addNewBottom();
            }
            // 设置底部边框的样式与左边边框相同
            bottomBorder.setVal(border.getVal());
            bottomBorder.setColor(border.getColor());
            bottomBorder.setSz(border.getSz());
            bottomBorder.setSpace(border.getSpace());
        }
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
        int tableMaxLineAllGridSpan = findTableMaxLineAllGridSpan(table);
        if (cellCount == 0) {
            XWPFTableCell tableCell = tableRow.addNewTableCell();
            CTTc ctTc = tableCell.getCTTc();
            CTTcPr ctTcPr = ctTc.isSetTcPr() ? ctTc.getTcPr() : ctTc.addNewTcPr();
            CTDecimalNumber gridSpan = ctTcPr.isSetGridSpan() ? ctTcPr.getGridSpan() : ctTcPr.addNewGridSpan();
            gridSpan.setVal(BigInteger.valueOf(tableMaxLineAllGridSpan));
        } else {
            if (toCol < fromCol || fromCol < 0 || toCol >= cellCount) {
                throw new IllegalArgumentException("col index out of bounds");
            }

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
            throw new RuntimeException(String.format("Total number of row is %d,The input row index(%d,%d) is incorrect", rows.size(), fromRow, toRow));
        }
        for (int i = fromRow; i <= toRow; i++) {
            XWPFTableRow row = rows.get(i);
            int size = row.getTableCells().size();
            size = size == 0 ? size : size - 1;
            mergeCellsHorizontal(table, i, row, 0, size);
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
