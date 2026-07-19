package com.deepoove.poi.plugin.table;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.exception.RenderException;
import com.deepoove.poi.render.compute.EnvModel;
import com.deepoove.poi.render.processor.EnvIterator;
import com.deepoove.poi.template.ElementTemplate;
import com.deepoove.poi.util.word.WordTableUtils;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlCursor;

import java.util.Collection;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

public class LoopCopyHeaderRowRenderPolicy extends AbstractLoopRowTableRenderPolicy {

    public LoopCopyHeaderRowRenderPolicy() {
        this(false);
    }

    public LoopCopyHeaderRowRenderPolicy(boolean onSameLine) {
        this("[", "]", onSameLine);
    }

    public LoopCopyHeaderRowRenderPolicy(String prefix, String suffix) {
        this(prefix, suffix, false);
    }

    public LoopCopyHeaderRowRenderPolicy(String prefix, String suffix, boolean onSameLine) {
        super();
        this.prefix = prefix;
        this.suffix = suffix;
        this.onSameLine = onSameLine;
    }

    public LoopCopyHeaderRowRenderPolicy(AbstractLoopRowTableRenderPolicy policy) {
        super(policy);
    }


    @Override
    public void render(ElementTemplate eleTemplate, Object data, XWPFTemplate template) {
        try {
            XWPFTableCell tagCell = this.dealPlaceTag(eleTemplate);
            int headerNumber = WordTableUtils.findCellVMergeNumber(tagCell);
            int templateRowIndex = this.getTemplateRowIndex(tagCell) + headerNumber - 1;
            int starRenderLocation = templateRowIndex;
            XWPFTable table = tagCell.getTableRow().getTable();

            int dataCount;
            if (data instanceof Collection) {
                dataCount = ((Collection<?>) data).size();
            } else {
                throw new RenderException("The data type is an " + data.getClass().getSimpleName() +
                    ", and the data type must be a collection");
            }

            Map<String, Object> globalEnv = template.getEnvModel().getEnv();
            Map<String, Object> original = new HashMap<>(globalEnv);
            int template_row_number = 1;
            int firstPageLine = 0;
            int pageLine = 0;
            int reduce = 0;
            boolean isRemoveNextLine = false;
            Object n = globalEnv.get(eleTemplate.getTagName() + "_number");
            int mode = 1;
            int writeCol = 0;
            boolean isDrawBorderOfFirstPage = false;
            try {
                if (n == null) {
                    // Subtract the default number of rows in the header by 1
                    pageLine = table.getRows().size() - 1;
                } else {
                    pageLine = Integer.parseInt(n.toString());
                }
                Object temp = globalEnv.get(eleTemplate.getTagName() + "_first_number");
                firstPageLine = temp != null ? Integer.parseInt(temp.toString()) : 0;
                temp = globalEnv.get(eleTemplate.getTagName() + "_mode");
                mode = temp != null ? Integer.parseInt(temp.toString()) : mode;
                temp = globalEnv.get(eleTemplate.getTagName() + "_write_col");
                writeCol = temp != null ? Integer.parseInt(temp.toString()) : writeCol;
                temp = globalEnv.get(eleTemplate.getTagName() + "_reduce");
                reduce = temp != null ? Integer.parseInt(temp.toString()) : reduce;
                temp = globalEnv.get(eleTemplate.getTagName() + "_remove_next_line");
                isRemoveNextLine = temp != null;
                temp = globalEnv.get(eleTemplate.getTagName() + "_fpdb");
                isDrawBorderOfFirstPage = temp != null;
            } catch (NumberFormatException ignore) {
            }

            // Delete blank XWPFParagraph after the table
            this.initDeal(template, globalEnv);
            WordTableUtils.removeLastBlankParagraph(xwpfDocument);

            this.setTemplateRowVMergeCol(table, templateRowIndex);

            Iterator<?> iterator = ((Iterable<?>) data).iterator();
            boolean hasNext = iterator.hasNext();
            int index = 0;
            boolean firstFlag = true;
            boolean firstPage = true;
            int allPage = countPageNumber(dataCount, 1, pageLine, firstPageLine);
            int currentPage = 1;
            XWPFTable nextTable = table;
            int templateRowIndex2 = templateRowIndex;
            int insertPosition;
            XWPFParagraph paragraph = null;
            while (hasNext) {
                Object root = iterator.next();
                hasNext = iterator.hasNext();

                firstPage = index < firstPageLine;
                if (index == 0 || index == firstPageLine || (index - firstPageLine) % pageLine == 0) {
                    if (index != 0) {
                        this.removeMultipleLine(template_row_number + (isRemoveNextLine ? 1 : 0),
                            table, templateRowIndex);
                    }
                    drawBottomBorder(currentPage, isDrawBorderOfFirstPage, table);
                    // 存在下一页，创建表格
                    table = nextTable;
                    if (currentPage <= allPage) {
                        // set page break
                        XmlCursor xmlCursor = table.getCTTbl().newCursor();
                        xmlCursor.toNextSibling();
                        paragraph = xwpfDocument.insertNewParagraph(xmlCursor);
                        WordTableUtils.setPageBreak(paragraph, 1);
                        WordTableUtils.setMinHeightParagraph(paragraph);
                        xmlCursor.toParent();
                        if (firstPage) {
                            xmlCursor.toNextSibling();
                            nextTable = xwpfDocument.insertNewTbl(xmlCursor);
                            nextTable.removeRow(0);
                            int rowIndex = WordTableUtils.findRowIndex(tagCell);
                            templateRowIndex2 = headerNumber;
                            int temp = 0;
                            for (int i = rowIndex; i < rowIndex + headerNumber + template_row_number; i++) {
                                WordTableUtils.copyLineContent(table.getRow(i), nextTable.insertNewTableRow(temp), temp++);
                            }
                            WordTableUtils.copyTableTblPr(table, nextTable);
                            nextTable.getCTTbl().setTblGrid(table.getCTTbl().getTblGrid());
                        } else {
                            nextTable = WordTableUtils.copyTable(xwpfDocument, table, xmlCursor);
                            templateRowIndex = templateRowIndex2;
                        }
                        xmlCursor.close();
                        firstFlag = true;
                        currentPage++;
                    }
                }

                insertPosition = templateRowIndex++;
                XWPFTableRow currentRow = table.getRow(insertPosition);
                if (!firstFlag) {
                    this.setVMerge(currentRow);
                } else {
                    firstFlag = false;
                }

                XWPFTableRow nextRow = table.insertNewTableRow(templateRowIndex);
                nextRow = WordTableUtils.copyLineContent(currentRow, nextRow, templateRowIndex);
                EnvIterator.makeEnv(globalEnv, ++index, index < dataCount);
                EnvModel.of(root, globalEnv);
                this.renderMultipleRow(table, insertPosition, insertPosition, resolver, documentProcessor);
                this.removeCurrentLineData(globalEnv, root);
                // 清除默认计算的缓存变量
                documentProcessor.clearElementProcessorInCache();
            }

            int insertLine;
            if (firstPage) {
                insertLine = firstPageLine - dataCount - reduce;
            } else if ((dataCount - firstPageLine) % pageLine == 0) {
                insertLine = 0;
            } else {
                insertLine = pageLine - (dataCount - firstPageLine) % pageLine - reduce;
            }
            this.fillBlankRow(insertLine, table, templateRowIndex);
            this.blankDeal(table, mode, templateRowIndex, insertLine, true, writeCol);

            if (paragraph != null) {
                WordTableUtils.removeParagraph(paragraph);
            }
            if (table != nextTable) {
                WordTableUtils.removeTable(xwpfDocument, nextTable);
            }
            this.removeMultipleLine(template_row_number + (isRemoveNextLine ? 1 : 0),
                table, templateRowIndex + insertLine);
            this.drawBottomBorder(currentPage, isDrawBorderOfFirstPage, table);
            globalEnv.putAll(original);
            afterloop(table, data);
        } catch (Exception e) {
            throw new RenderException("HackLoopTable for " + eleTemplate + " error: " + e.getMessage(), e);
        }
    }

    protected void afterloop(XWPFTable table, Object data) {
    }
}
