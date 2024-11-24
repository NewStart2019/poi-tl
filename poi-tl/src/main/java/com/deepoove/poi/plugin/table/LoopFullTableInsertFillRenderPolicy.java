package com.deepoove.poi.plugin.table;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.exception.RenderException;
import com.deepoove.poi.policy.RenderPolicy;
import com.deepoove.poi.render.compute.EnvModel;
import com.deepoove.poi.render.processor.EnvIterator;
import com.deepoove.poi.template.ElementTemplate;
import com.deepoove.poi.util.WordTableUtils;
import com.deepoove.poi.xwpf.NiceXWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlCursor;

import java.util.Collection;
import java.util.HashMap;
import java.util.Map;

public class LoopFullTableInsertFillRenderPolicy extends AbstractLoopRowTableRenderPolicy implements RenderPolicy {

    public LoopFullTableInsertFillRenderPolicy() {
        this(false);
    }

    public LoopFullTableInsertFillRenderPolicy(boolean onSameLine) {
        this("[", "]", onSameLine);
    }

    public LoopFullTableInsertFillRenderPolicy(String prefix, String suffix) {
        this(prefix, suffix, false);
    }

    public LoopFullTableInsertFillRenderPolicy(String prefix, String suffix, boolean onSameLine) {
        this.prefix = prefix;
        this.suffix = suffix;
        this.onSameLine = onSameLine;
    }

    public LoopFullTableInsertFillRenderPolicy(AbstractLoopRowTableRenderPolicy policy) {
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
            this.initDeal(template, globalEnv);
            Map<String, Object> original = new HashMap<>(globalEnv);

            int templateRowNumber = 1;
            int pageLine = 0;
            int reduce = 0;
            boolean isRemoveNextLine = false;
            Object n = globalEnv.get(eleTemplate.getTagName() + "_number");
            int mode = 1;
            boolean isFill = true;
            boolean isVMerge = false;
            try {
                if (n == null) {
                    // Subtract the default number of rows in the header by 1
                    pageLine = table.getRows().size() - 1;
                } else {
                    pageLine = Integer.parseInt(n.toString());
                }
                Object temp = globalEnv.get(eleTemplate.getTagName() + "_mode");
                mode = temp != null ? Integer.parseInt(temp.toString()) : mode;
                temp = globalEnv.get(eleTemplate.getTagName() + "_reduce");
                reduce = temp != null ? Integer.parseInt(temp.toString()) : reduce;
                temp = globalEnv.get(eleTemplate.getTagName() + "_remove_next_line");
                isRemoveNextLine = temp != null;
                temp = globalEnv.get(eleTemplate.getTagName() + "_row_number");
                templateRowNumber = temp == null ? templateRowNumber : Integer.parseInt(temp.toString());
                temp = globalEnv.get(eleTemplate.getTagName() + "_nofill");
                isFill = temp == null;
                temp = globalEnv.get(eleTemplate.getTagName() + "_vmerge");
                isVMerge = temp != null;
            } catch (NumberFormatException ignore) {
            }
            if (pageLine % templateRowNumber != 0) {
                throw new RenderException("The size of each page should be a multiple of the number of lines in the template for multi line rendering!");
            }

            int index = 0;
            XWPFTable nextTable = table;
            int tempTemplateRowIndex = templateRowIndex;
            int insertPosition;
            int tableCount = countPageNumber(dataCount, templateRowNumber, pageLine, pageLine);
            int perPageNumber = pageLine / templateRowNumber;
            int currentTableIndex = 1;
            boolean firstFlag = true;

            XWPFParagraph paragraph = null;
            NiceXWPFDocument xwpfDocument = template.getXWPFDocument();
            for (Object root : (Iterable<?>) data) {
                // Determine whether to cross page and copy a new table across pages
                if (index % perPageNumber == 0) {
                    if (index != 0) {
                        this.removeMultipleLine(templateRowNumber, table, tempTemplateRowIndex);
                        if (isRemoveNextLine && table.getRows().size() > tempTemplateRowIndex) {
                            table.removeRow(tempTemplateRowIndex);
                        }
                        this.renderMultipleRow(table, tempTemplateRowIndex, -1, resolver, documentProcessor);
                    }
                    table = nextTable;
                    if (currentTableIndex <= tableCount) {
                        // set page break
                        XmlCursor xmlCursor = table.getCTTbl().newCursor();
                        xmlCursor.toNextSibling();
                        paragraph = xwpfDocument.insertNewParagraph(xmlCursor);
                        WordTableUtils.setPageBreak(paragraph, 1);
                        WordTableUtils.setMinHeightParagraph(paragraph);

                        xmlCursor.toParent();
                        nextTable = WordTableUtils.copyTable(xwpfDocument, table, xmlCursor);
                        xmlCursor.close();
                        tempTemplateRowIndex = templateRowIndex;
                        currentTableIndex++;
                        firstFlag = true;
                    }
                }

                // Insert new rows into the original table
                EnvIterator.makeEnv(globalEnv, ++index, index < dataCount);
                EnvModel.of(root, globalEnv);
                insertPosition = tempTemplateRowIndex;
                tempTemplateRowIndex += templateRowNumber;
                for (int i = 0; i < templateRowNumber; i++) {
                    int currentIndex = insertPosition + i;
                    XWPFTableRow currentRow = table.getRow(currentIndex);
                    XWPFTableRow nextRow = table.insertNewTableRow(tempTemplateRowIndex + i);
                    nextRow = WordTableUtils.copyLineContent(currentRow, nextRow, tempTemplateRowIndex + i);
                    if (isVMerge) {
                        if (!firstFlag) {
                            this.setVMerge(currentRow);
                        } else {
                            firstFlag = false;
                        }
                    }
                    this.renderMultipleRow(table, currentIndex, currentIndex, resolver, documentProcessor);
                }
                this.removeCurrentLineData(globalEnv, root);
            }

            if (isFill && dataCount > 0) {
                int insertLine;
                if (currentTableIndex == 2) {
                    insertLine = pageLine - dataCount * templateRowNumber - reduce;
                } else if (dataCount % perPageNumber == 0) {
                    insertLine = 0;
                } else {
                    insertLine = pageLine - dataCount % perPageNumber * templateRowNumber - reduce;
                }
                this.fillBlankRow(insertLine, table, tempTemplateRowIndex);
                this.blankDeal(table, mode, tempTemplateRowIndex, insertLine);
                tempTemplateRowIndex += insertLine;
                if (paragraph != null) {
                    WordTableUtils.removeParagraph(paragraph);
                }
                if (table != nextTable) {
                    WordTableUtils.removeTable(xwpfDocument, nextTable);
                }
                globalEnv.clear();
                globalEnv.putAll(original);
            } else {
                if (paragraph != null) {
                    WordTableUtils.removeParagraph(paragraph);
                }
                if (table != nextTable) {
                    WordTableUtils.removeTable(xwpfDocument, nextTable);
                }
            }
            this.removeMultipleLine(templateRowNumber, table, tempTemplateRowIndex);
            if (isRemoveNextLine) {
                table.removeRow(tempTemplateRowIndex);
            }
            this.renderMultipleRow(table, tempTemplateRowIndex, -1, resolver, documentProcessor);
            afterloop(table, data);
        } catch (Exception e) {
            throw new RenderException("HackLoopTable for " + eleTemplate + " error: " + e.getMessage(), e);
        }
    }

    protected void afterloop(XWPFTable table, Object data) {
    }
}
