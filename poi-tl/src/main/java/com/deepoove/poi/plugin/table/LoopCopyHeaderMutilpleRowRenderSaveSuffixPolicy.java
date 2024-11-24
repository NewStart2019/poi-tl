package com.deepoove.poi.plugin.table;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.exception.RenderException;
import com.deepoove.poi.policy.RenderPolicy;
import com.deepoove.poi.render.compute.EnvModel;
import com.deepoove.poi.render.processor.EnvIterator;
import com.deepoove.poi.template.ElementTemplate;
import com.deepoove.poi.template.MetaTemplate;
import com.deepoove.poi.util.WordTableUtils;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlCursor;

import java.util.*;

public class LoopCopyHeaderMutilpleRowRenderSaveSuffixPolicy extends AbstractLoopRowTableRenderPolicy implements RenderPolicy {

    public LoopCopyHeaderMutilpleRowRenderSaveSuffixPolicy() {
        this(false);
    }

    public LoopCopyHeaderMutilpleRowRenderSaveSuffixPolicy(boolean onSameLine) {
        this("[", "]", onSameLine);
    }

    public LoopCopyHeaderMutilpleRowRenderSaveSuffixPolicy(String prefix, String suffix) {
        this(prefix, suffix, false);
    }

    public LoopCopyHeaderMutilpleRowRenderSaveSuffixPolicy(String prefix, String suffix, boolean onSameLine) {
        this.prefix = prefix;
        this.suffix = suffix;
        this.onSameLine = onSameLine;
    }

    public LoopCopyHeaderMutilpleRowRenderSaveSuffixPolicy(AbstractLoopRowTableRenderPolicy policy) {
        super(policy);
    }

    @Override
    public void render(ElementTemplate eleTemplate, Object data, XWPFTemplate template) {
        try {
            XWPFTableCell tagCell = this.dealPlaceTag(eleTemplate);
            int headerNumber = WordTableUtils.findCellVMergeNumber(tagCell);
            int templateRowIndex = this.getTemplateRowIndex(tagCell) + headerNumber - 1;
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
            Object n = globalEnv.get(eleTemplate.getTagName() + "_number");
            int mode = 1;
            boolean isDrawBorderOfFirstPage = false;
            int tableExternalFooterLine = 0;
            boolean isFill = true;
            try {
                if (n == null) {
                    // Subtract the default number of rows in the header by 1
                    pageLine = table.getRows().size() - 1;
                } else {
                    pageLine = Integer.parseInt(n.toString());
                }
                Object temp = globalEnv.get(eleTemplate.getTagName() + "_row_number");
                template_row_number = temp == null ? template_row_number : Integer.parseInt(temp.toString());
                temp = globalEnv.get(eleTemplate.getTagName() + "_first_number");
                firstPageLine = temp != null ? Integer.parseInt(temp.toString()) : firstPageLine;
                temp = globalEnv.get(eleTemplate.getTagName() + "_mode");
                mode = temp != null ? Integer.parseInt(temp.toString()) : mode;
                temp = globalEnv.get(eleTemplate.getTagName() + "_reduce");
                reduce = temp != null ? Integer.parseInt(temp.toString()) : reduce;
                temp = globalEnv.get(eleTemplate.getTagName() + "_fpdb");
                isDrawBorderOfFirstPage = temp != null;
                temp = globalEnv.get(eleTemplate.getTagName() + "_external_footer");
                tableExternalFooterLine = temp != null ? Integer.parseInt(temp.toString()) : tableExternalFooterLine;
                temp = globalEnv.get(eleTemplate.getTagName() + "_nofill");
                isFill = temp == null;
            } catch (NumberFormatException ignore) {
            }
            if (template_row_number > firstPageLine) {
                throw new RenderException("Template rendering with more lines than the first page is not supported!");
            }
            if (firstPageLine % template_row_number != 0 || pageLine % template_row_number != 0) {
                throw new RenderException("The size of each page should be a multiple of the number of lines in the template for multi line rendering!");
            }

            this.initDeal(template, globalEnv);
            // Delete blank XWPFParagraph after the table
            WordTableUtils.removeLastBlankParagraph(xwpfDocument);
            Iterator<?> iterator = ((Iterable<?>) data).iterator();
            boolean hasNext = iterator.hasNext();
            int index = 0;
            boolean firstPage = true;
            int allPage = countPageNumber(dataCount, template_row_number, pageLine, firstPageLine);
            int firstNumber = firstPageLine / template_row_number;
            int perPageNumber = pageLine / template_row_number;
            int currentPage = 1;
            XWPFTable nextTable = table;
            int templateRowIndex2 = templateRowIndex;
            int insertPosition;
            XWPFParagraph paragraph = null;
            while (hasNext) {
                Object root = iterator.next();
                hasNext = iterator.hasNext();

                firstPage = index < firstNumber;
                if (index == 0 || index == firstNumber || ((index > firstNumber) && (index - firstNumber) % perPageNumber == 0)) {
                    if (index != 0) {
                        removeMultipleLine(template_row_number, table, templateRowIndex);
                    }
                    // Set the bottom border of the table to the left border style
                    drawBottomBorder(currentPage, isDrawBorderOfFirstPage, table);
                    table = nextTable;
                    if (currentPage <= allPage) {
                        // set page break
                        XmlCursor xmlCursor = table.getCTTbl().newCursor();
                        xmlCursor.toNextSibling();
                        paragraph = xwpfDocument.insertNewParagraph(xmlCursor);
                        WordTableUtils.setPageBreak(paragraph, 1);
                        WordTableUtils.setMinHeightParagraph(paragraph);
                        if (firstPage) {
                            xmlCursor.toParent();
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
                            xmlCursor.toParent();
                            nextTable = WordTableUtils.copyTable(xwpfDocument, table, xmlCursor);
                            templateRowIndex = templateRowIndex2;
                        }
                        xmlCursor.close();
                        currentPage++;
                    }
                }

                insertPosition = templateRowIndex;
                templateRowIndex += template_row_number;
                EnvIterator.makeEnv(globalEnv, ++index, index < dataCount);
                EnvModel.of(root, globalEnv);
                for (int i = 0; i < template_row_number; i++) {
                    XWPFTableRow currentRow = table.getRow(insertPosition + i);
                    XWPFTableRow nextRow = table.insertNewTableRow(templateRowIndex + i);
                    nextRow = WordTableUtils.copyLineContent(currentRow, nextRow, templateRowIndex + i);
                    currentRow.getTableCells().forEach(cell -> {
                        List<MetaTemplate> templates = resolver.resolveBodyElements(cell.getBodyElements());
                        documentProcessor.process(templates);
                    });
                }
                removeCurrentLineData(globalEnv, root);
            }

            if (isFill) {
                int insertLine;
                if (firstPage) {
                    insertLine = firstPageLine - index * template_row_number;
                    // You can finish writing the remaining part at the end of the first page
                    if (insertLine >= tableExternalFooterLine) {
                        insertLine = insertLine - tableExternalFooterLine - reduce;
                        this.fillBlankRow(insertLine, table, templateRowIndex);
                        this.removeMultipleLine(template_row_number, table, templateRowIndex + insertLine);
                        if (paragraph != null) {
                            WordTableUtils.removeParagraph(paragraph);
                        }
                        if (table != nextTable) {
                            WordTableUtils.removeTable(xwpfDocument, nextTable);
                        }
                        this.blankDeal(table, mode, templateRowIndex, insertLine, true);
                        this.drawBottomBorder(currentPage, isDrawBorderOfFirstPage, table);
                    } else {
                        // The first table filling section
                        this.fillBlankRow(insertLine, table, templateRowIndex);
                        this.blankDeal(table, mode, templateRowIndex, insertLine, true);
                        this.removeMultipleLine(template_row_number, table, templateRowIndex + insertLine);
                        boolean isNoRemain = insertLine == 0;
                        // The second table filling section
                        insertLine = pageLine - tableExternalFooterLine - reduce;
                        this.fillBlankRow(insertLine, nextTable, headerNumber);
                        this.removeMultipleLine(template_row_number, nextTable, headerNumber + insertLine);
                        this.blankDeal(nextTable, mode, headerNumber, insertLine, isNoRemain);
                        this.drawBottomBorder(currentPage, isDrawBorderOfFirstPage, nextTable);
                    }
                } else if ((dataCount - firstNumber) % perPageNumber == 0) {
                    this.removeMultipleLine(template_row_number, table, templateRowIndex);

                    insertLine = pageLine - (dataCount - firstNumber) % perPageNumber * template_row_number - tableExternalFooterLine - reduce;
                    this.fillBlankRow(insertLine, nextTable, headerNumber);
                    this.removeMultipleLine(template_row_number, nextTable, headerNumber + insertLine);
                    this.blankDeal(nextTable, mode, headerNumber, insertLine, true);
                    this.drawBottomBorder(currentPage, isDrawBorderOfFirstPage, nextTable);
                } else if ((perPageNumber - (dataCount - firstNumber) % perPageNumber) * template_row_number >= tableExternalFooterLine) {
                    insertLine = pageLine - (dataCount - firstNumber) % perPageNumber * template_row_number - tableExternalFooterLine - reduce;
                    this.fillBlankRow(insertLine, table, templateRowIndex);
                    this.removeMultipleLine(template_row_number, table, templateRowIndex + insertLine);
                    if (paragraph != null) {
                        WordTableUtils.removeParagraph(paragraph);
                    }
                    if (table != nextTable) {
                        WordTableUtils.removeTable(xwpfDocument, nextTable);
                    }
                    this.blankDeal(table, mode, templateRowIndex, insertLine, true);
                    this.drawBottomBorder(currentPage, isDrawBorderOfFirstPage, table);
                } else {
                    insertLine = pageLine - (dataCount - firstNumber) % perPageNumber * template_row_number;
                    // The first table filling section
                    this.fillBlankRow(insertLine, table, templateRowIndex);
                    this.blankDeal(table, mode, templateRowIndex, insertLine, true);
                    removeMultipleLine(template_row_number, table, templateRowIndex + insertLine);
                    boolean isNoRemain = insertLine == 0;
                    // The second table filling section
                    insertLine = pageLine - tableExternalFooterLine - reduce;
                    this.fillBlankRow(insertLine, nextTable, headerNumber);
                    this.removeMultipleLine(template_row_number, nextTable, headerNumber + insertLine);
                    this.blankDeal(nextTable, mode, headerNumber, insertLine, isNoRemain);
                    this.drawBottomBorder(currentPage, isDrawBorderOfFirstPage, nextTable);
                }
            } else {
                this.removeMultipleLine(template_row_number, table, templateRowIndex);
                if (paragraph != null) {
                    WordTableUtils.removeParagraph(paragraph);
                }
                if (table != nextTable) {
                    WordTableUtils.removeTable(xwpfDocument, nextTable);
                }
                this.drawBottomBorder(currentPage, isDrawBorderOfFirstPage, table);
            }
            afterloop(table, data);
            globalEnv.putAll(original);
        } catch (Exception e) {
            throw new RenderException("HackLoopTable for " + eleTemplate + " error: " + e.getMessage(), e);
        }
    }

    protected void afterloop(XWPFTable table, Object data) {
    }
}

