package com.deepoove.poi.util.word;

import com.deepoove.poi.util.ReflectionUtils;
import com.deepoove.poi.util.TableTools;
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
import org.openxmlformats.schemas.drawingml.x2006.main.CTBlip;
import org.openxmlformats.schemas.drawingml.x2006.picture.CTPicture;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.w3c.dom.Node;

import java.awt.*;
import java.awt.font.FontRenderContext;
import java.awt.geom.AffineTransform;
import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.List;

/**
 * copy → clean -> remove → find → set → merge → move
 * TODO update doc property id， DrawingSupport.updateDocPrId(table);
 */
@SuppressWarnings("unused")
public class WordTableUtils {

    private static final Logger logger = LoggerFactory.getLogger(WordTableUtils.class);

    /**
     * 无垂直合并标记。
     */
    private static final int VMERGE_NONE = 0;
    /**
     * 垂直合并的后续单元格（vMerge=continue，含未显式写 val 的情况）。
     */
    private static final int VMERGE_CONTINUE = 1;
    /**
     * 垂直合并的起始单元格（vMerge=restart）。
     */
    private static final int VMERGE_RESTART = 2;

    /**
     * 1 点对应的 twips 数（1 point = 20 twips）。
     */
    private static final int TWIPS_PER_POINT = 20;

    /** 查找行内所有图片 blip 节点的 XPath（用于跨文档复制后重映射图片关系）。 */
    private static final String XPATH_BLIP =
        "declare namespace a='http://schemas.openxmlformats.org/drawingml/2006/main' .//a:blip";

    public static XWPFTable copyTable(XWPFDocument doc, XWPFTable sourceTable) {
        return copyTable(doc, sourceTable, null, true, false);
    }

    public static XWPFTable copyTable(XWPFDocument doc, XWPFTable sourceTable, boolean isTail) {
        return copyTable(doc, sourceTable, null, true, isTail);
    }

    public static XWPFTable copyTable(XWPFDocument doc, XWPFTable sourceTable, XmlCursor xmlCursor) {
        return copyTable(doc, sourceTable, xmlCursor, true, false);
    }

    public static XWPFTable copyToXmlCursorBefore(XWPFDocument doc, XWPFTable sourceTable, XmlCursor xmlCursor) {
        return copyTable(doc, sourceTable, xmlCursor, false, false);
    }

    /**
     * <p>Copy the specified table after the specified {@link XmlCursor xmlCursor}. If <b>{@link XmlCursor xmlCursor}</b> parameter is <b>empty</b>,
     * it will be added by default before the position of the <b>current</b> tbale.</p>
     * <p>If isTail is True, the new table will be placed at the end of the document.</p>
     * <p><b>Notice: </b></p>
     * <ul>
     *     <li>If the {@link XmlCursor xmlCursor} parameter is created externally, remember to release the resource yourself</li>
     *     <li>The copy table function can only be used to copy within the same table</li>
     * </ul>
     *
     * @param doc         {@link XWPFDocument doc}
     * @param sourceTable {@link XWPFTable sourceTable}
     * @param xmlCursor   {@link XmlCursor xmlCursor}
     * @param isAfter     boolean, whether to place the new table after the current table
     * @param isTail      boolean, whether to place the new table at the end of the document
     * @return {@link XWPFTable}
     */
    @SuppressWarnings("unchecked")
    public static XWPFTable copyTable(XWPFDocument doc, XWPFTable sourceTable, XmlCursor xmlCursor, boolean isAfter, boolean isTail) {
        if (doc == null || sourceTable == null) {
            throw new RuntimeException("The parameters passed in cannot be empty!");
        }
        // 记录游标是否由本方法内部创建：外部传入的游标按约定由调用方负责释放，
        // 内部创建的游标必须在 finally 中关闭，否则会一直占用 XML store 的定位资源。
        boolean selfCreatedCursor = (xmlCursor == null);
        if (selfCreatedCursor) {
            xmlCursor = sourceTable.getCTTbl().newCursor();
        }
        try {
            // isAfter 时移到源表的下一个兄弟节点；若源表已是最后一个元素，toNextSibling 返回
            // false 且游标仍停留在源表上，insertNewTbl 会把新表插到源表之后，结果仍正确。
            if (isAfter) {
                xmlCursor.toNextSibling();
            }
            if (isTail) {
                while (xmlCursor.toNextSibling()) ;
            }
            XWPFTable table = doc.insertNewTbl(xmlCursor);
            table.removeRow(0);
            CTTbl ctTbl = table.getCTTbl();
            ctTbl.set(sourceTable.getCTTbl());
            CTRow[] trArray = ctTbl.getTrArray();
            List<XWPFTableRow> tableRows =
                (List<XWPFTableRow>) ReflectionUtils.getValue("tableRows", table);
            for (int i = 0; i < trArray.length; i++) {
                XWPFTableRow row = new XWPFTableRow(trArray[i], table);
                tableRows.add(i, row);
            }
            return table;
        } finally {
            if (selfCreatedCursor) {
                xmlCursor.close();
            }
        }
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
        // 源表没有可复制的 tblPr 时保持目标不变，避免 set(null) 静默清空目标已有的表样式。
        if (sourceTblPr == null) {
            return;
        }
        CTTbl targetCTTbl = targetTable.getCTTbl();
        CTTblPr targetTblPr = targetCTTbl.getTblPr();
        if (targetTblPr == null) {
            // 目标缺少 tblPr 时按 schema 位置新建，避免旧实现对 null 调用 set 抛出 NPE。
            targetTblPr = targetCTTbl.addNewTblPr();
        }
        targetTblPr.set(sourceTblPr);
    }

    /**
     * 复制表格行（含行内所有单元格内容与样式）到目标行。
     *
     * <p>支持三种场景：同一文档的同一表格、同一文档的不同表格、不同文档的不同表格。
     * 行级属性（trPr，如行高/对齐）、单元格属性（TcPr，如 gridSpan/vMerge/宽高/边框）与段落、
     * run 样式都会随 XML 一并复制；跨文档复制时，图片关系（r:id）会被重映射到目标文档，
     * 避免出现悬空引用。</p>
     * <p>提示：注意原来的表格行的单元格有跨列需压迫自行处理这种特殊情况</p>
     *
     * @param currentLine    源行
     * @param nextLine       目标行（已存在于目标表格中）
     * @param isIncludeStyle true：连同样式完整复制；false：仅复制文本，保留目标行原有样式
     */
    public static void copyLine(XWPFTableRow currentLine, XWPFTableRow nextLine, boolean isIncludeStyle) {
        if (currentLine == null || nextLine == null) {
            return;
        }
        XWPFTable targetTable = nextLine.getTable();
        XWPFDocument sourceDoc = currentLine.getTable().getBody().getXWPFDocument();
        XWPFDocument targetDoc = targetTable.getBody().getXWPFDocument();
        boolean sameDoc = sourceDoc == targetDoc;

        if (isIncludeStyle) {
            // 深拷贝整段行 XML：trPr、各 TcPr（gridSpan/vMerge/宽高/边框）与全部内容，
            // 同表复制时源行与目标行底层对象互不影响。
            nextLine.getCtRow().set(currentLine.getCtRow());
            // 去掉目标行的跨行（vMerge）设置，使复制出的行成为独立行，不再与源行共享垂直合并。
            if (!sameDoc) {
                // 跨文档：XML 中的图片 r:id 仍指向源文档，需要把图片数据加入目标文档并重映射。
                remapRowPictureRelations(nextLine, sourceDoc, targetDoc);
            }
        } else {
            // 仅复制文本内容，保留目标行自身的 trPr 与单元格 TcPr。
            copyLineContentOnly(currentLine, nextLine);
        }

        // 底层 CTRow 变化后，同步刷新目标行内部的单元格包装，避免高层 getCell 与底层 tc 脱节。
        refreshRowCells(nextLine, targetTable);
    }

    /**
     * 跨文档复制后，把目标行内所有图片的 blip 引用重映射到目标文档。
     *
     * @param targetRow 目标行（已复制源行 XML）
     * @param sourceDoc 源文档
     * @param targetDoc 目标文档
     */
    private static void remapRowPictureRelations(XWPFTableRow targetRow, XWPFDocument sourceDoc,
                                                 XWPFDocument targetDoc) {
        XmlObject[] blips = targetRow.getCtRow().selectPath(XPATH_BLIP);
        for (XmlObject obj : blips) {
            CTBlip blip = (CTBlip) obj;
            if (!blip.isSetEmbed()) {
                continue;
            }
            String oldId = blip.getEmbed();
            XWPFPictureData picture = sourceDoc.getPictureDataByID(oldId);
            if (picture == null) {
                continue;
            }
            try {
                String newId = targetDoc.addPictureData(picture.getData(), picture.getPictureType());
                blip.setEmbed(newId);
            } catch (InvalidFormatException e) {
                logger.warn("Failed to remap picture relation {} when copying row across documents", oldId, e);
            }
        }
    }

    /**
     * Only copy the text content to the target row, retaining the style and structure of each cell in the target row.
     */
    private static void copyLineContentOnly(XWPFTableRow source, XWPFTableRow target) {
        List<XWPFTableCell> sourceCells = source.getTableCells();
        for (int i = 0; i < sourceCells.size(); i++) {
            XWPFTableCell targetCell = i < target.getTableCells().size()
                ? target.getCell(i) : target.addNewTableCell();
            // copyCellContent 保留目标段落属性，仅复制源段落文本。
            copyCellContent(sourceCells.get(i), targetCell, false);
        }
    }


    /**
     * 底层 CTRow 变化后，原地重建目标行的单元格包装列表，保证高层 API 与底层 tc 一致。
     */
    @SuppressWarnings("unchecked")
    private static void refreshRowCells(XWPFTableRow row, XWPFTable table) {
        List<XWPFTableCell> cells =
            (List<XWPFTableCell>) ReflectionUtils.getValue("tableCells", row);
        cells.clear();
        for (CTTc ctTc : row.getCtRow().getTcList()) {
            cells.add(new XWPFTableCell(ctTc, row, table.getBody()));
        }
    }

    /**
     * Copy the content of the current line to the next line. If the next line is a newly added line,
     * then directly copy the entire XML of the current line to the next line. Otherwise, just copy
     * the content to the next line.
     * <p><b>Tips:</b> </p>
     * <li>Cross table copying of the same file has not been tested, please use with caution</li>
     * <li>Cross table replication of different files has not been tested, please use with caution</li>
     *
     * @param currentLine      current line
     * @param nextLine         next line
     * @param templateRowIndex next line row index
     * @return next line
     */
    public static XWPFTableRow copyLineContent(XWPFTableRow currentLine, XWPFTableRow nextLine, int templateRowIndex) {
        if (currentLine == null || nextLine == null) {
            return nextLine;
        }
        XWPFTable table = nextLine.getTable();
        if (CollectionUtils.isEmpty(nextLine.getTableCells())) {
            // 目标是无单元格的新建空行：直接在底层 CTRow 上深拷贝源行 XML（无需 XmlCursor， 再原地刷新单元格包装。
            nextLine.getCtRow().set(currentLine.getCtRow());
            refreshRowCells(nextLine, table);
            return nextLine;
        }

        // 目标已有单元格：逐格复制内容，单元格不足时新建。
        List<XWPFTableCell> sourceCells = currentLine.getTableCells();
        int nextCellSize = nextLine.getTableCells().size() - 1;
        for (int i = 0; i < sourceCells.size(); i++) {
            boolean isNoHasCell = nextCellSize < i;
            XWPFTableCell nextCell = isNoHasCell ? nextLine.addNewTableCell() : nextLine.getCell(i);
            copyCellContent(sourceCells.get(i), nextCell, isNoHasCell);
        }
        // 复制行级属性（行高/对齐等）
        copyTrPr(currentLine.getCtRow(), nextLine.getCtRow());
        return nextLine;
    }

    /**
     * 把源行的 trPr（行高、对齐、分页控制等）复制到目标行。
     *
     * <p>源行没有 trPr 时保持目标不变；目标缺少 trPr 时按 schema 位置新建，避免空指针。
     */
    private static void copyTrPr(CTRow source, CTRow target) {
        if (!source.isSetTrPr()) {
            return;
        }
        CTTrPr targetTrPr = target.isSetTrPr() ? target.getTrPr() : target.addNewTrPr();
        targetTrPr.set(source.getTrPr());
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
        if (isIncludeStyle) {
            target.getCTP().setPPr(source.getCTPPr());
        }
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

    /**
     * @param templateRow
     */
    public static void cleanRowTextContent(XWPFTableRow templateRow) {
        List<XWPFTableCell> tableCells = templateRow.getTableCells();
        tableCells.forEach(cell -> {
            List<XWPFParagraph> paragraphs = cell.getParagraphs();
            for (int i = paragraphs.size() - 1; i >= 0; i--) {
                cleanParagraphContent(paragraphs.get(i));
            }
        });
    }

    /**
     * 清空行内所有段落的文本内容，但保留段落属性与单元格结构（TcPr 中的 gridSpan、vMerge、宽度等）。
     *
     * @param row 待清空的底层行（通常为深拷贝出的副本）
     */
    public static void cleanRowTextKeepStructure(CTRow row) {
        for (CTTc tc : row.getTcList()) {
            for (CTP p : tc.getPList()) {
                Node paragraphNode = p.getDomNode();
                List<Node> toRemove = new ArrayList<>();
                Node child = paragraphNode.getFirstChild();
                while (child != null) {
                    // 仅保留段落属性 pPr，其余元素子节点（run、超链接等）全部移除。
                    if (child.getNodeType() == Node.ELEMENT_NODE
                        && !"pPr".equals(child.getLocalName())) {
                        toRemove.add(child);
                    }
                    child = child.getNextSibling();
                }
                for (Node node : toRemove) {
                    paragraphNode.removeChild(node);
                }
            }
        }
    }

    public static void cleanCellContent(XWPFTableCell cell) {
        removeAllParagraphsOfCell(cell);
    }

    /**
     * Clear the content of XWPFParagraph, but does not remove the XWPFRun of paragraph.
     *
     * @param paragraph {@link XWPFParagraph paragraph}
     */
    public static void cleanParagraphContent(XWPFParagraph paragraph) {
        if (paragraph == null) {
            return;
        }
        List<XWPFRun> runs = paragraph.getRuns();
        for (int i = runs.size() - 1; i >= 0; i--) {
            XWPFRun xwpfRun = runs.get(i);

            CTR ctr = xwpfRun.getCTR();
            cleanXWPFRunContent(xwpfRun);
            // 添加一个空的 <w:t> 节点，防止后续操作出错（可选）
            CTText text = ctr.addNewT();
            text.setStringValue("");
        }
    }

    /**
     * Clean up XWPFRun content
     *
     * @param run {@link XWPFRun run}
     */
    public static void cleanXWPFRunContent(XWPFRun run) {
        if (run == null) {
            return;
        }
        CTR ctr = run.getCTR();
        if (ctr == null) {
            return;
        }

        // Remove all sub elements such as<w:t>,<w: br/>,<w: tab/>, etc
        for (int temp = ctr.sizeOfTArray() - 1; temp >= 0; temp--) {
            ctr.removeT(temp);
        }
        for (int temp = ctr.sizeOfBrArray() - 1; temp >= 0; temp--) {
            ctr.removeBr(temp);
        }
        for (int temp = ctr.sizeOfTabArray() - 1; temp >= 0; temp--) {
            ctr.removeTab(temp);
        }
        for (int temp = ctr.sizeOfContentPartArray() - 1; temp >= 0; temp--) {
            ctr.removeContentPart(temp);
        }
        for (int temp = ctr.sizeOfDelTextArray() - 1; temp >= 0; temp--) {
            ctr.removeDelText(temp);
        }
        for (int temp = ctr.sizeOfInstrTextArray() - 1; temp >= 0; temp--) {
            ctr.removeInstrText(temp);
        }
        for (int temp = ctr.sizeOfDelInstrTextArray() - 1; temp >= 0; temp--) {
            ctr.removeDelInstrText(temp);
        }
        for (int temp = ctr.sizeOfNoBreakHyphenArray() - 1; temp >= 0; temp--) {
            ctr.removeNoBreakHyphen(temp);
        }
        for (int temp = ctr.sizeOfSoftHyphenArray() - 1; temp >= 0; temp--) {
            ctr.removeSoftHyphen(temp);
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

    /**
     * Remove row self.
     * <p>The deleted rows do not consider the issue of cross-row and column deletion processing. Please use with caution</p>
     *
     * @param row {@link XWPFTableRow row}
     */
    public static void removeRow(XWPFTableRow row) {
        if (row == null) {
            return;
        }
        XWPFTable table = row.getTable();
        int rowIndex = findRowIndex(table, row.getCtRow());
        table.removeRow(rowIndex);
    }

    public static void removeRowVMerge(XWPFTableRow tableRow) {
        removeRowVMerge(tableRow.getCtRow());
    }

    /**
     * 移除行内所有单元格的垂直合并（vMerge，即跨行）标记，保留其余 TcPr（宽度、边框、gridSpan 等）。
     *
     * <p>无论 vMerge 是 restart 还是 continue（含未显式写 val 的情况）都移除，使复制出的行
     * 不再与目标表中的相邻行发生垂直合并。
     */
    public static void removeRowVMerge(CTRow ctRow) {
        for (CTTc ctTc : ctRow.getTcList()) {
            if (ctTc.isSetTcPr()) {
                CTTcPr tcPr = ctTc.getTcPr();
                if (tcPr.isSetVMerge()) {
                    tcPr.unsetVMerge();
                }
            }
        }
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
        // 注意：不能因为段落没有 run 就跳过删除。分页段落通过段落属性 w:pageBreakBefore
        // 实现（见 XWPFParagraph.setPageBreak(true)），本身不包含任何 run，若在此提前
        // 返回会导致末位分页符无法删除，进而把表后的备注挤到下一页。
        IBody body = paragraph.getBody();
        if (body instanceof XWPFDocument) {
            XWPFDocument document = (XWPFDocument) body;
            int posOfTable = document.getPosOfParagraph(paragraph);
            if (posOfTable < 0) {
                return;
            }
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
                        // Do not remove paragraphs if there are images or text present
                        if (StringUtils.isNotBlank(run.text()) || CollectionUtils.isNotEmpty(run.getEmbeddedPictures())) {
                            isEmpty = false;
                            break;
                        }
                    }
                    if (isEmpty) {
                        xwpfDocument.removeBodyElement(i);
                    } else {
                        break;
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
     * <p>Get the number of cross columns.</p>
     * <p>Note: The number of cross columns is 1 by default</p>
     * <p> If the input parameter is empty, return 0 </p>
     *
     * @param cell {@link XWPFTableCell cell}
     * @return int
     */
    public static int findTableCellGridSpan(XWPFTableCell cell) {
        if (cell == null) {
            return 0;
        }
        CTTc ctTc = cell.getCTTc();
        if (!ctTc.isSetTcPr()) {
            return 1;
        }
        CTTcPr tcPr = ctTc.getTcPr();
        if (!tcPr.isSetGridSpan()) {
            return 1;
        } else {
            CTDecimalNumber gridSpan = tcPr.getGridSpan();
            return gridSpan.getVal().intValue();
        }
    }

    /**
     * Query the specified row, specify the column position, including the number of cross columns
     *
     * @param table    {@link XWPFTable table}
     * @param startRow row index
     * @param colIndex col index
     * @return int
     */
    public static int findHorizontalMergedCells(XWPFTable table, int startRow, int colIndex) {
        if (table == null) {
            return 0;
        }
        List<XWPFTableRow> rows = table.getRows();
        if (startRow < 0 || startRow >= rows.size()) {
            return 0;
        }
        List<XWPFTableCell> tableCells = rows.get(startRow).getTableCells();
        if (colIndex < 0 || colIndex >= tableCells.size()) {
            return 0;
        }
        int count = 0;
        for (int i = 0; i < tableCells.size(); i++) {
            XWPFTableCell cell = tableCells.get(i);
            count += findTableCellGridSpan(cell);
            if (i == colIndex) {
                break;
            }
        }
        return count;
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
        // Calculate the number of cross columns in a specified cell
        int realColNumber = findHorizontalMergedCells(table, startRow, colIndex);
        int i = startRow + 1;
        int size = table.getRows().size();
        for (; i < size; i++) {
            if (table.getRow(i).getCell(colIndex) == null) {
                break;
            }
            List<XWPFTableCell> tableCells = table.getRow(i).getTableCells();
            int tempRealColNumber = 0;
            // If the corresponding position is found without crossing columns, end the loop
            boolean flag = false;
            for (int col = 0; col < tableCells.size(); col++) {
                XWPFTableCell cell = tableCells.get(col);
                tempRealColNumber += findTableCellGridSpan(cell);
                // Same number of columns
                if (tempRealColNumber == realColNumber) {
                    CTTc ctTc = cell.getCTTc();
                    if (!ctTc.isSetTcPr()) {
                        flag = true;
                        break;
                    }
                    CTTcPr tcPr = ctTc.getTcPr();
                    if (!tcPr.isSetVMerge()) {
                        flag = true;
                        break;
                    }
                    CTVMerge vMerge = tcPr.getVMerge();
                    if (vMerge == null || vMerge.getVal() == STMerge.RESTART) {
                        flag = true;
                        break;
                    }
                }
            }
            if (flag) {
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

    public static XWPFTableCell findMaxWidthCellInRow(XWPFTableRow row) {
        if (row == null) {
            return null;
        }
        int maxWidth = 0;
        List<XWPFTableCell> tableCells = row.getTableCells();
        XWPFTableCell tableCell = tableCells.get(0);
        for (XWPFTableCell tempCell : tableCells) {
            if (tableCell.getWidth() < tempCell.getWidth()) {
                tableCell = tempCell;
            }
        }
        return tableCell;
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

    /**
     * 以底层 CTRow 对象身份在 CTTbl 中定位真实当前索引。
     *
     * @param table 表格
     * @param ctRow 目标底层行
     * @return 当前索引；不存在时返回 -1
     */
    public static int findRowIndex(XWPFTable table, CTRow ctRow) {
        CTRow[] rows = table.getCTTbl().getTrArray();
        for (int i = 0; i < rows.length; i++) {
            if (rows[i] == ctRow) {
                return i;
            }
        }
        return -1;
    }

    public static int findRowHeight(XWPFTableRow row) {
        return row.getHeight();
    }


    /**
     * 获取单元格宽度（单位：twips），支持水平合并（gridSpan/跨列）。
     *
     * <p>优先按表格网格（tblGrid）累加该单元格覆盖的所有网格列宽，普通单元格（span=1）与
     * 跨列单元格口径一致；网格缺失或为自动宽度（列宽为 0）时回退到单元格自身的 tcW，
     * 合并单元格通常已在其中写入合并后的总宽。
     *
     * @param cell 单元格
     * @return 宽度（twips）；无法确定时返回 0
     */
    public static int findCellWidth(XWPFTableCell cell) {
        if (cell == null) {
            return 0;
        }
        XWPFTable table = cell.getTableRow().getTable();
        int startGridCol = findCellStartGridCol(cell);
        int gridSpan = findTableCellGridSpan(cell);

        // 首选：按 tblGrid 累加 gridSpan 覆盖的网格列。
        CTTblGrid grid = table.getCTTbl().getTblGrid();
        if (grid != null && grid.sizeOfGridColArray() > 0) {
            int gridCount = grid.sizeOfGridColArray();
            int sum = 0;
            boolean covered = true;
            for (int i = 0; i < gridSpan; i++) {
                int colIndex = startGridCol + i;
                if (colIndex >= gridCount) {
                    covered = false;
                    break;
                }
                sum += toTwips(grid.getGridColArray(colIndex).getW());
            }
            if (covered && sum > 0) {
                return sum;
            }
        }

        // 回退：单元格自身 tcW（旧实现用 (int) 强转 BigInteger 会抛 ClassCastException，这里统一走 toTwips）。
        CTTcPr tcPr = cell.getCTTc().getTcPr();
        if (tcPr != null && tcPr.isSetTcW()) {
            return toTwips(tcPr.getTcW().getW());
        }
        return 0;
    }

    /**
     * 获取单元格高度（单位：twips），支持垂直合并（vMerge/跨行）。
     *
     * <p>若单元格处于垂直合并中，先定位到合并起始行（vMerge=restart），再累加起始行及其后
     * 同网格列上所有 continuation 行的行高；普通单元格直接返回所在行高度。
     *
     * @param cell 单元格
     * @return 高度（twips）；无法确定时返回 0
     */
    public static int findCellHeight(XWPFTableCell cell) {
        if (cell == null) {
            return 0;
        }
        XWPFTable table = cell.getTableRow().getTable();
        int gridCol = findCellStartGridCol(cell);

        // 若传入的是 continue，则向上回溯到合并起始行（restart）。
        int startRow = findRowIndex(cell);
        while (startRow > 0
            && vMergeState(findCellAtGridColumn(table.getRow(startRow), gridCol)) == VMERGE_CONTINUE) {
            startRow--;
        }

        int totalHeight = 0;
        int rowCount = table.getRows().size();
        for (int r = startRow; r < rowCount; r++) {
            XWPFTableCell current = findCellAtGridColumn(table.getRow(r), gridCol);
            if (current == null) {
                break;
            }
            int state = vMergeState(current);
            if (r == startRow) {
                // 起始行：计入行高；若并非合并单元格（无 vMerge）则直接结束。
                totalHeight += findRowHeight(table.getRow(r));
                if (state == VMERGE_NONE) {
                    break;
                }
            } else if (state == VMERGE_CONTINUE) {
                totalHeight += findRowHeight(table.getRow(r));
            } else {
                // 遇到新的 restart 或无 vMerge，当前垂直合并结束。
                break;
            }
        }
        return totalHeight;
    }

    /**
     * 计算单元格在表格网格中的起始列索引（累加其前面所有单元格的 gridSpan）。
     *
     * @param cell 单元格
     * @return 起始网格列索引（从 0 开始）
     */
    public static int findCellStartGridCol(XWPFTableCell cell) {
        CTTc target = cell.getCTTc();
        int startGridCol = 0;
        for (XWPFTableCell sibling : cell.getTableRow().getTableCells()) {
            if (sibling.getCTTc() == target) {
                break;
            }
            startGridCol += findTableCellGridSpan(sibling);
        }
        return startGridCol;
    }

    /**
     * 按网格列索引定位某行中覆盖该列的单元格（兼容行内 gridSpan 导致的列错位）。
     *
     * @param row     行
     * @param gridCol 网格列索引
     * @return 覆盖该列的单元格；行内不存在时返回 null
     */
    public static XWPFTableCell findCellAtGridColumn(XWPFTableRow row, int gridCol) {
        int accumulated = 0;
        for (XWPFTableCell candidate : row.getTableCells()) {
            int span = findTableCellGridSpan(candidate);
            if (gridCol < accumulated + span) {
                return candidate;
            }
            accumulated += span;
        }
        return null;
    }

    /**
     * 读取单元格的垂直合并状态。
     *
     * <p>vMerge 未显式写 val 时按 OOXML 语义默认 continue。
     *
     * @param cell 单元格
     * @return {@link #VMERGE_NONE}（无合并）、{@link #VMERGE_CONTINUE}（后续）或
     * {@link #VMERGE_RESTART}（起始）
     */
    public static int vMergeState(XWPFTableCell cell) {
        if (cell == null) {
            return VMERGE_NONE;
        }
        CTTcPr tcPr = cell.getCTTc().getTcPr();
        if (tcPr == null || !tcPr.isSetVMerge()) {
            return VMERGE_NONE;
        }
        CTVMerge vMerge = tcPr.getVMerge();
        if (vMerge != null && vMerge.isSetVal() && vMerge.getVal() == STMerge.RESTART) {
            return VMERGE_RESTART;
        }
        return VMERGE_CONTINUE;
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
     * @return Integer | null | 0 则表示没有跨行
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

    /**
     * <p>Get the font size of the run</p>
     * <p><b>Tips:</b></p>
     * <p>&nbsp;&nbsp;&nbsp;&nbsp;If {@link XWPFRun run} is empty, return 0</p>
     * <p>&nbsp;&nbsp;&nbsp;&nbsp;If font size is not set, return 0</p>
     *
     * @param run {@link XWPFRun run}
     * @return font size
     */
    public static double findFontSize(XWPFRun run) {
        if (run == null) {
            return 0;
        }
        // 优先使用常规字号（w:sz）。此前在常规字号非空时反而用复杂脚本字号覆盖，
        // 当仅设置 w:sz（无 szCs）时会被覆盖成 null 而返回 0，导致上层误判为默认字号。
        Double fontSizeAsDouble = run.getFontSizeAsDouble();
        if (fontSizeAsDouble != null) {
            return fontSizeAsDouble;
        }
        // 常规字号缺失时，回退到复杂脚本字号（w:szCs）。
        Double complexScriptSize = run.getComplexScriptFontSizeAsDouble();
        return complexScriptSize == null ? 0 : complexScriptSize;
    }

    /**
     * <p>Move a body element (paragraph or table) to the specified position in the document</p>
     * <p>Synchronously updates the bodyElements list, category lists (paragraphs/tables), and the underlying XML structure</p>
     * <p>Uses XmlCursor navigated from CTBody by index to avoid attribute-region positioning errors</p>
     *
     * @param document    {@link XWPFDocument doc}
     * @param bodyElement {@link IBodyElement bodyElement} element to move
     * @param position    target position (0-based) in the bodyElements list
     */
    @SuppressWarnings("unchecked")
    public static void setElementPosition(XWPFDocument document, IBodyElement bodyElement, int position) {
        if (document == null || bodyElement == null) {
            logger.warn("document or bodyElement is null");
            return;
        }
        List<IBodyElement> bodyElements = (List<IBodyElement>) ReflectionUtils.getValue("bodyElements", document);
        int sourceIndex = bodyElements.indexOf(bodyElement);
        if (sourceIndex < 0) {
            logger.warn("Body element not found in document");
            return;
        }
        if (position < 0 || position > bodyElements.size()) {
            throw new IndexOutOfBoundsException("Position " + position + " is out of range [0, " + bodyElements.size() + "]");
        }
        if (sourceIndex == position) {
            return;
        }

        CTBody body = document.getDocument().getBody();
        XmlCursor srcCursor = null;
        XmlCursor targetCursor = null;
        try {
            // ── 1. Position source cursor at the STARTELEMENT of the source element ──
            srcCursor = body.newCursor();
            srcCursor.toFirstChild(); // enter first child element
            for (int i = 0; i < sourceIndex; i++) {
                srcCursor.toNextSibling();
            }

            // ── 2. Position target cursor at the insertion point ──
            targetCursor = body.newCursor();
            if (position == bodyElements.size()) {
                // Insert at end — move cursor to just before CTBody's ENDELEMENT
                targetCursor.toEndToken();
            } else {
                // Insert before the element at target index
                targetCursor.toFirstChild();
                for (int i = 0; i < position; i++) {
                    targetCursor.toNextSibling();
                }
            }

            // ── 3. Move the entire element via XmlCursor (XMLBeans native) ──
            srcCursor.moveXml(targetCursor);

            // ── 4. Update bodyElements list ──
            bodyElements.remove(sourceIndex);
            int actualTarget = sourceIndex < position ? position - 1 : position;
            bodyElements.add(actualTarget, bodyElement);

            // ── 5. Rebuild category lists to keep consistent order ──
            rebuildCategoryLists(document, bodyElements);
        } finally {
            if (srcCursor != null) {
                srcCursor.close();
            }
            if (targetCursor != null) {
                targetCursor.close();
            }
        }
    }

    /**
     * Rebuild paragraphs and tables category lists from bodyElements to preserve document order
     */
    @SuppressWarnings("unchecked")
    private static void rebuildCategoryLists(XWPFDocument document, List<IBodyElement> bodyElements) {
        List<Paragraph> paragraphs = (List<Paragraph>) ReflectionUtils.getValue("paragraphs", document);
        List<XWPFTable> tables = (List<XWPFTable>) ReflectionUtils.getValue("tables", document);
        paragraphs.clear();
        tables.clear();
        for (IBodyElement element : bodyElements) {
            if (element.getElementType() == BodyElementType.PARAGRAPH) {
                paragraphs.add((XWPFParagraph) element);
            } else if (element.getElementType() == BodyElementType.TABLE) {
                tables.add((XWPFTable) element);
            }
        }
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

    /**
     * <p>Move a table row from source index to target index</p>
     * <p>Both Java object list and underlying CT XML structure are updated synchronously</p>
     *
     * @param table       {@link XWPFTable table}
     * @param sourceIndex source row index (0-based)
     * @param targetIndex target row index (0-based), the row will be inserted before this position
     */
    @SuppressWarnings("unchecked")
    public static void moveRowTo(XWPFTable table, int sourceIndex, int targetIndex) {
        if (table == null) {
            logger.warn("table is null");
            return;
        }
        List<XWPFTableRow> rows = (List<XWPFTableRow>) ReflectionUtils.getValue("tableRows", table);
        int size = rows.size();
        if (sourceIndex < 0 || sourceIndex >= size) {
            logger.warn("sourceIndex {} is out of range (0-{})", sourceIndex, size - 1);
            return;
        }
        if (targetIndex < 0 || targetIndex > size) {
            logger.warn("targetIndex {} is out of range (0-{})", targetIndex, size);
            return;
        }
        if (sourceIndex == targetIndex) {
            return;
        }
        // Remove from source position
        XWPFTableRow sourceRow = rows.remove(sourceIndex);
        // Adjust target index if source was before target
        int actualTarget = sourceIndex < targetIndex ? targetIndex - 1 : targetIndex;
        // Insert at target position
        rows.add(actualTarget, sourceRow);

        // Sync with underlying CT XML structure
        CTTbl ctTbl = table.getCTTbl();
        CTRow ctRow = sourceRow.getCtRow();
        // Remove from source position in CT array
        ctTbl.removeTr(sourceIndex);
        // Insert at target position in CT array
        ctTbl.insertNewTr(actualTarget);
        ctTbl.setTrArray(actualTarget, ctRow);
    }

    /**
     * Move a row up by one position
     *
     * @param table    {@link XWPFTable table}
     * @param rowIndex current row index
     */
    public static void moveRowUp(XWPFTable table, int rowIndex) {
        moveRowTo(table, rowIndex, rowIndex - 1);
    }

    /**
     * Move a row down by one position
     *
     * @param table    {@link XWPFTable table}
     * @param rowIndex current row index
     */
    public static void moveRowDown(XWPFTable table, int rowIndex) {
        moveRowTo(table, rowIndex, rowIndex + 1);
    }

    /**
     * Move a specified row object to target index
     *
     * @param row         {@link XWPFTableRow row}
     * @param targetIndex target position
     */
    public static void moveRow(XWPFTableRow row, int targetIndex) {
        if (row == null) {
            logger.warn("row is null");
            return;
        }
        XWPFTable table = row.getTable();
        int sourceIndex = findRowIndex(row);
        if (sourceIndex < 0) {
            logger.warn("row not found in table");
            return;
        }
        moveRowTo(table, sourceIndex, targetIndex);
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
     * @param type        {@link STHeightRule.Enum type} , default is {@link STHeightRule.Enum EXACT}: indicates a fixed height
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
        height.setHRule(type);
    }

    public static void setCellWidth(XWPFTableCell cell, String width) {
        CTTcPr tcPr = cell.getCTTc().isSetTcPr() ? cell.getCTTc().getTcPr() : cell.getCTTc().addNewTcPr();
        CTTblWidth tblWidth = tcPr.isSetTcW() ? tcPr.getTcW() : tcPr.addNewTcW();
        tblWidth.setType(STTblWidth.DXA);
        tblWidth.setW(new BigInteger(width));
    }

    public static void setCellWidth(XWPFTableCell cell, int width) {
        CTTblWidth tblWidth = CTTblWidth.Factory.newInstance();
        tblWidth.setType(STTblWidth.DXA);
        tblWidth.setW(BigInteger.valueOf(width));
        CTTcPr tcPr = TableTools.getTcPr(cell);
        tcPr.setTcW(tblWidth);
    }

    /**
     * <p>Set vertical merge</p>
     *
     * @param cell  {@link XWPFTableCell cell}
     * @param merge {@link STMerge.Enum merge}
     */
    public static void setCellVmerge(XWPFTableCell cell, STMerge.Enum merge) {
        CTTcPr tcPr = TableTools.getTcPr(cell);
        CTVMerge vMerge = tcPr.isSetVMerge() ? tcPr.getVMerge() : tcPr.addNewVMerge();
        vMerge.setVal(merge);
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
        CTBorder ctBorder = ctTcBorders.isSetTr2Bl() ? ctTcBorders.getTr2Bl() : ctTcBorders.addNewTr2Bl();
        ctBorder.setVal(stBorder);
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
     * Create a section break before the specified label, with <b>xmlCursor</b> is blank and added at the end by default
     *
     * @param doc       {@link XWPFDocument doc}  required
     * @param type      {@link STSectionMark.Enum type} required
     * @param xmlCursor {@link XmlCursor xmlCursor}
     */
    public static void setSectionBreak(XWPFDocument doc, STSectionMark.Enum type, XmlCursor xmlCursor) {
        if (doc == null || type == null) {
            return;
        }
        XWPFParagraph paragraph = xmlCursor == null ? doc.createParagraph() : doc.insertNewParagraph(xmlCursor);
        CTSectType ctSectType = CTSectType.Factory.newInstance();
        ctSectType.setVal(type);

        CTP ctp = paragraph.getCTP();
        CTPPr ctpPr = ctp.isSetPPr() ? ctp.getPPr() : ctp.addNewPPr();
        CTSectPr ctSectPr = ctpPr.isSetSectPr() ? ctpPr.getSectPr() : ctpPr.addNewSectPr();
        ctSectPr.setType(ctSectType);
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

        // Set global border style
        CTTblPr tblPr = table.getCTTbl().getTblPr();
        if (tblPr == null) tblPr = table.getCTTbl().addNewTblPr();
        CTTblBorders ctTblBorders = tblPr.isSetTblBorders() ? tblPr.getTblBorders() : tblPr.addNewTblBorders();
        CTBorder globalBottomBorder = ctTblBorders.isSetBottom() ? ctTblBorders.getBottom() : ctTblBorders.addNewBottom();
        globalBottomBorder.setVal(border.getVal());
        globalBottomBorder.setColor(border.getColor());
        globalBottomBorder.setSz(border.getSz());
        globalBottomBorder.setSpace(border.getSpace());

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

            for (int colIndex = toCol + 1; colIndex > fromCol; colIndex--) {
                XWPFTableCell cell = tableRow.getCell(colIndex); // 注意：每次删一个，索引会变
                if (cell != null) {
                    tableRow.getCtRow().removeTc(colIndex);
                    tableRow.getTableCells().remove(fromCol + 1);
                }
            }
        }
    }

    public static void unMergeCells(XWPFTableRow row, int startCellIndex, boolean isAddSpan) {
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
     * cancel cell merge by specified cell
     *
     * @param cellRow    {@link XWPFTableCell cell}
     * @param isUnhmerge whether to cancel the horizontal merge
     * @param isUnvmerge whether to cancel the vertical merge
     */
    public static void umMergeCell(XWPFTableCell cellRow, boolean isUnhmerge, boolean isUnvmerge) {
        if (cellRow == null) {
            return;
        }
        CTTcPr tcPr = cellRow.getCTTc().getTcPr();
        if (tcPr == null) {
            return;
        }
        if (tcPr.isSetHMerge() && isUnhmerge) {
            tcPr.unsetHMerge();
        }
        if (tcPr.isSetGridSpan() && isUnhmerge) {
            CTDecimalNumber gridSpan = tcPr.getGridSpan();
            if (gridSpan != null && gridSpan.getVal() != null) {
                int gridSpanValue = gridSpan.getVal().intValue();
                if (gridSpanValue > 1) {
                    tcPr.unsetGridSpan();
                }
            }
        }

        if (tcPr.isSetVMerge() && isUnvmerge) {
            tcPr.unsetVMerge();
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
    public static void mergeMultipleLine(XWPFTable table, int fromRow, int toRow) {
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

    /**
     * 默认字号（点）。
     */
    private static final double DEFAULT_FONT_SIZE = 12;
    /**
     * 最小可读字号（点）。
     */
    private static final double MIN_FONT_SIZE = 1;
    /**
     * 字号递减步长（点）。
     */
    private static final double FONT_SIZE_STEP = 0.5;
    /**
     * 高度安全系数：对计算出的所需高度预留 8% 余量。
     *
     * <p>不同渲染器对同一字体的行高口径不同（宋体 win 指标为 1.0× em，而 hhea/typo 含行间距
     * 为 1.14× em，LibreOffice 取后者、Word 取前者），且半磅取整、居中对齐都会累积误差。
     * 预留固定比例余量可保证选出的字号在各渲染器下都不被 exact 固定行高裁剪，属工程上的
     * 保守取值，而非随机魔数。
     */
    private static final double HEIGHT_SAFETY_FACTOR = 1.08;

    /**
     * 根据单元格宽高与文本内容计算可容纳的字号（单位：点）。
     *
     * <p>计算思路：按当前字号用 AWT 测量文本宽度（点 × 20 换算为 twips），结合可用宽度得到
     * 占用行数，再结合单行高度与段前段后间距得到总高度；若总高度超过单元格高度，则以 0.5 点为
     * 步长逐级缩小，直到放得下或达到最小可读下限。
     *
     * <p>本版本修复：
     * <ol>
     *     <li>补充 cell/content 空值与非正宽高防护，避免 NPE 与除零退化；</li>
     *     <li>粗体、斜体使用位或组合，两者可同时生效，并通过 {@link XWPFRun#isBold()}、
     *     {@link XWPFRun#isItalic()} 正确识别 {@code w:val=0} 的显式关闭；</li>
     *     <li>段前段后优先使用 twips 间距（before/after），百分之一行单位（beforeLines）
     *     按当前行高换算，不再当作 twips 直接相减；</li>
     *     <li>行距根据 {@code lineRule} 区分：EXACT/AT_LEAST 为 twips，AUTO 为 1/240 行倍数；</li>
     *     <li>使用 {@link Font#deriveFont(float)} 支持小数字号，测量与返回值保持一致。</li>
     * </ol>
     *
     * @param cell       目标单元格
     * @param content    待渲染文本
     * @param cellWidth  单元格宽度（twips）
     * @param cellHeight 单元格高度（twips）
     * @return 计算出的字号（点）；无法缩放时返回默认或最小字号
     */
    public static double adjustFontSize(XWPFTableCell cell, String content, double cellWidth, double cellHeight) {
        if (cell == null || cell.getParagraphs().isEmpty()) {
            return DEFAULT_FONT_SIZE;
        }
        XWPFParagraph paragraph = cell.getParagraphs().get(0);
        if (paragraph.getRuns().isEmpty()) {
            return DEFAULT_FONT_SIZE;
        }
        XWPFRun run = paragraph.getRuns().get(0);

        // 基准字号：取 run 实际字号，缺失或非法时回退默认字号。
        double baseSize = findFontSize(run);
        if (baseSize <= 0) {
            baseSize = DEFAULT_FONT_SIZE;
        }
        // 空文本在任意字号下都放得下，直接返回基准字号。
        if (content == null || content.isEmpty()) {
            return baseSize;
        }
        // 宽高非法时不做缩放，避免除零产生 0.5 之类的退化结果。
        if (cellWidth <= 0 || cellHeight <= 0) {
            return baseSize;
        }

        // 单元格内边距（tcMar → 表级 tblCellMar → Word 默认值）。文字只能在扣除内边距后的
        // 区域内排版，不扣除会高估可用宽度，导致字号偏大、文字被单元格裁剪。
        Insets cellMargin = cellMarginTwips(cell);

        // 可用宽度：扣除单元格左右内边距与段落左右缩进；结果非正时回退到单元格宽度。
        double usableWidth = cellWidth - cellMargin.left - cellMargin.right
            - paragraph.getIndentationLeft() - paragraph.getIndentationRight();
        if (usableWidth <= 0) {
            usableWidth = cellWidth;
        }
        // 可用高度：扣除单元格上下内边距；结果非正时回退到单元格高度。
        double usableHeight = cellHeight - cellMargin.top - cellMargin.bottom;
        if (usableHeight <= 0) {
            usableHeight = cellHeight;
        }
        // 首行额外缩进（首行缩进 - 悬挂），计入文本宽度。
        int firstLineExtra = paragraph.getFirstLineIndent();

        CTSpacing spacing = readSpacing(paragraph);

        int style = Font.PLAIN;
        // 位或组合，粗体与斜体可同时生效；isBold/isItalic 已正确处理 val=0 的显式关闭。
        if (run.isBold()) {
            style |= Font.BOLD;
        }
        if (run.isItalic()) {
            style |= Font.ITALIC;
        }

        FontRenderContext frc = new FontRenderContext(new AffineTransform(), true, true);
        // 选择与文本实际脚本匹配的测量字体：中文等东亚文字必须用 eastAsia 字体，
        // 若用 ascii 的 Times New Roman 会明显低估宽度、算出偏大的字号。
        Font templateFont = resolveMeasureFont(run, content, style);

        double candidate = baseSize;
        while (candidate >= MIN_FONT_SIZE) {
            Font font = templateFont.deriveFont((float) candidate);
            // 文本宽度（twips）：点 × 20，再加上首行额外缩进。
            double textWidth = font.getStringBounds(content, frc).getWidth() * TWIPS_PER_POINT
                + firstLineExtra;
            int lineCount = (int) Math.ceil(textWidth / usableWidth);
            double lineHeight = lineHeightTwips(spacing, font, frc);
            double neededHeight = lineCount * lineHeight + verticalPaddingTwips(spacing, lineHeight);
            // 预留安全余量后再比较，确保居中块在各渲染器行高口径与取整误差下都不越过行格边界。
            if (neededHeight * HEIGHT_SAFETY_FACTOR <= usableHeight) {
                return candidate;
            }
            candidate -= FONT_SIZE_STEP;
        }
        // 达到最小可读下限仍放不下：返回最小字号，避免返回不可见的 0.5 以下值。
        return MIN_FONT_SIZE;
    }

    /**
     * 读取单元格内边距（twips）。优先取单元格级 w:tcMar，其次表级 w:tblCellMar，
     * 两者都未设置时回退到 Word 默认值（左右各 108、上下 0）。
     *
     * @return 按 top/left/bottom/right 顺序封装的内边距
     */
    private static Insets cellMarginTwips(XWPFTableCell cell) {
        int top = 0;
        /**
         *  Word 默认单元格左右内边距（twips，各 108）。
         */
        int DEFAULT_CELL_MARGIN_X = 108;
        int left = DEFAULT_CELL_MARGIN_X;
        int bottom = 0;
        int right = DEFAULT_CELL_MARGIN_X;

        CTTcMar tcMar = null;
        CTTcPr tcPr = cell.getCTTc().getTcPr();
        if (tcPr != null && tcPr.isSetTcMar()) {
            tcMar = tcPr.getTcMar();
        }
        CTTblCellMar tblMar = null;
        CTTblPr tblPr = cell.getTableRow().getTable().getCTTbl().getTblPr();
        if (tblPr != null && tblPr.isSetTblCellMar()) {
            tblMar = tblPr.getTblCellMar();
        }

        if (tcMar != null) {
            top = marginTwips(tcMar.getTop(), top);
            left = marginTwips(tcMar.getLeft(), left);
            bottom = marginTwips(tcMar.getBottom(), bottom);
            right = marginTwips(tcMar.getRight(), right);
        } else if (tblMar != null) {
            top = marginTwips(tblMar.getTop(), top);
            left = marginTwips(tblMar.getLeft(), left);
            bottom = marginTwips(tblMar.getBottom(), bottom);
            right = marginTwips(tblMar.getRight(), right);
        }
        return new Insets(top, left, bottom, right);
    }

    /**
     * 读取单个内边距（仅支持绝对 dxa 单位），缺失或非法时返回默认值。
     */
    private static int marginTwips(CTTblWidth width, int fallback) {
        if (width == null || width.getW() == null) {
            return fallback;
        }
        Object w = width.getW();
        if (w instanceof Number) {
            return ((Number) w).intValue();
        }
        try {
            return Integer.parseInt(w.toString());
        } catch (NumberFormatException e) {
            return fallback;
        }
    }

    /**
     * 选择用于测量文本宽度的 AWT 字体。
     *
     * <p>Word 中 ascii/hAnsi 与 eastAsia 分别使用不同字体：中文等东亚文字实际由 eastAsia
     * 字体渲染。若一律用 ascii 字体（如 Times New Roman）测量，其东亚字符回退字形明显偏窄，
     * 会低估文本宽度、算出偏大的字号。这里按文本脚本优先选择对应字体，并通过
     * {@link Font#canDisplayUpTo(String)} 确认字体可完整渲染，否则逐级回退，最后使用
     * 逻辑字体 {@link Font#SANS_SERIF}（Windows 上映射为含中文字形的字体）。
     *
     * @return 字号为 1 的基准字体，后续用 {@link Font#deriveFont(float)} 应用候选字号
     */
    private static Font resolveMeasureFont(XWPFRun run, String content, int style) {
        String ascii = run.getFontFamily(XWPFRun.FontCharRange.ascii);
        String hAnsi = run.getFontFamily(XWPFRun.FontCharRange.hAnsi);
        String eastAsia = run.getFontFamily(XWPFRun.FontCharRange.eastAsia);

        // 按内容脚本排列候选字体优先级：含东亚文字时 eastAsia 优先，否则 ascii 优先。
        List<String> families = new ArrayList<>();
        if (containsEastAsian(content)) {
            addIfNotBlank(families, eastAsia);
            addIfNotBlank(families, ascii);
            addIfNotBlank(families, hAnsi);
        } else {
            addIfNotBlank(families, ascii);
            addIfNotBlank(families, hAnsi);
            addIfNotBlank(families, eastAsia);
        }

        for (String family : families) {
            Font font = new Font(family, style, 1);
            // 字体名可被解析且能完整显示全部文本时才采用。
            if (font.canDisplayUpTo(content) == -1) {
                return font;
            }
        }
        // 兜底：逻辑 SansSerif 在 Windows 上可覆盖中文。
        return new Font(Font.SANS_SERIF, style, 1);
    }

    private static void addIfNotBlank(List<String> list, String value) {
        if (value != null && !value.trim().isEmpty() && !list.contains(value)) {
            list.add(value);
        }
    }

    /**
     * 判断文本是否包含东亚文字（CJK 统一表意文字及扩展、日文假名、韩文谚文、CJK 标点、
     * 全角字符等），用于决定测量字体的优先级。
     */
    private static boolean containsEastAsian(String content) {
        for (int i = 0; i < content.length(); ) {
            int cp = content.codePointAt(i);
            i += Character.charCount(cp);
            if (isEastAsianCodePoint(cp)) {
                return true;
            }
        }
        return false;
    }

    private static boolean isEastAsianCodePoint(int cp) {
        return (cp >= 0x4E00 && cp <= 0x9FFF)     // CJK 统一表意文字
            || (cp >= 0x3400 && cp <= 0x4DBF)     // CJK 扩展 A
            || (cp >= 0x20000 && cp <= 0x2A6DF)   // CJK 扩展 B
            || (cp >= 0x3040 && cp <= 0x309F)     // 平假名
            || (cp >= 0x30A0 && cp <= 0x30FF)     // 片假名
            || (cp >= 0xAC00 && cp <= 0xD7AF)     // 韩文谚文音节
            || (cp >= 0x3000 && cp <= 0x303F)     // CJK 标点
            || (cp >= 0xFF00 && cp <= 0xFFEF);    // 全角 ASCII/标点
    }

    /**
     * 读取段落的间距设置，未设置时返回 null。
     */
    private static CTSpacing readSpacing(XWPFParagraph paragraph) {
        CTP ctp = paragraph.getCTP();
        if (ctp.isSetPPr() && ctp.getPPr().isSetSpacing()) {
            return ctp.getPPr().getSpacing();
        }
        return null;
    }

    /**
     * 计算单行高度（twips）。
     *
     * <p>EXACT 行距直接取固定 twips；AT_LEAST 取固定值与自然行高的较大者（避免自然行高
     * 超过下限而被低估）；AUTO（多倍行距）按 1/240 行倍数乘以自然行高；未设置行距时
     * 直接使用自然行高。
     *
     * <p>自然行高取自 AWT 字体的 ascent + descent + leading（含东亚字体的内部行距），
     * 而非简单的字号 em 尺寸——宋体等字体的单倍行高约为 em 的 1.14 倍，用 em 会低估行高、
     * 导致在固定行高（exact）单元格内文字被裁剪。
     */
    private static double lineHeightTwips(CTSpacing spacing, Font font, FontRenderContext frc) {
        double naturalHeight = naturalLineHeightTwips(font, frc);
        if (spacing == null || !spacing.isSetLine()) {
            return naturalHeight;
        }
        int line = toTwips(spacing.getLine());
        STLineSpacingRule.Enum rule = spacing.isSetLineRule()
            ? spacing.getLineRule() : STLineSpacingRule.AUTO;
        if (rule == STLineSpacingRule.EXACT) {
            return line;
        }
        if (rule == STLineSpacingRule.AT_LEAST) {
            return Math.max(line, naturalHeight);
        }
        // AUTO：line 以 1/240 行表示，240 即单倍自然行高。
        return line / 240.0 * naturalHeight;
    }

    /**
     * 测量字体的自然单倍行高（twips）。样本同时包含拉丁与 CJK 字符，确保取到
     * 东亚字体完整的 ascent/descent/leading。
     */
    private static double naturalLineHeightTwips(Font font, FontRenderContext frc) {
        java.awt.font.LineMetrics lm = font.getLineMetrics("Ay中", frc);
        double linePt = lm.getAscent() + lm.getDescent() + lm.getLeading();
        return linePt * TWIPS_PER_POINT;
    }

    /**
     * 计算段前段后占用的垂直间距（twips）。
     *
     * <p>优先使用 twips 间距 before/after；仅设置百分之一行单位 beforeLines/afterLines 时，
     * 按当前行高换算。
     */
    private static double verticalPaddingTwips(CTSpacing spacing, double lineHeightTwips) {
        if (spacing == null) {
            return 0;
        }
        double padding = 0;
        if (spacing.isSetBefore()) {
            padding += toTwips(spacing.getBefore());
        } else if (spacing.isSetBeforeLines()) {
            padding += spacing.getBeforeLines().doubleValue() / 100.0 * lineHeightTwips;
        }
        if (spacing.isSetAfter()) {
            padding += toTwips(spacing.getAfter());
        } else if (spacing.isSetAfterLines()) {
            padding += spacing.getAfterLines().doubleValue() / 100.0 * lineHeightTwips;
        }
        return padding;
    }

    /**
     * 将间距值安全转换为 twips（兼容 Number 与字符串表示）。
     */
    private static int toTwips(Object value) {
        if (value == null) {
            return 0;
        }
        if (value instanceof Number) {
            return ((Number) value).intValue();
        }
        return Integer.parseInt(value.toString());
    }


    /**
     * Migrate the content and cell attributes of the "restart" cell to the corresponding column in the specified reserved row.
     *
     * @param table           {@link XWPFTable}
     * @param restartRowIndex restart the line
     * @param col             column index
     * @param targetRowIndex  first reserved row index
     */
    public static void moveRestartCell(XWPFTable table, int restartRowIndex, int col, int targetRowIndex) {
        XWPFTableCell source = table.getRow(restartRowIndex).getCell(col);
        XWPFTableRow targetRow = table.getRow(targetRowIndex);
        XWPFTableCell target = targetRow.getCell(col);
        if (target == null) {
            target = targetRow.createCell();
        }
        // 复制段落内容及 TcPr（含 restart 的 vMerge 标记）。
        WordTableUtils.copyVerticallyCellContent(source, target);
        WordTableUtils.cleanCellContent(source);
    }
}
