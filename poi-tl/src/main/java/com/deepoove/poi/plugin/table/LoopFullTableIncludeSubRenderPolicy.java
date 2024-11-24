package com.deepoove.poi.plugin.table;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.exception.RenderException;
import com.deepoove.poi.policy.RenderPolicy;
import com.deepoove.poi.render.compute.EnvModel;
import com.deepoove.poi.render.processor.EnvIterator;
import com.deepoove.poi.template.ElementTemplate;
import com.deepoove.poi.util.WordTableUtils;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlCursor;

import java.util.Collection;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

public class LoopFullTableIncludeSubRenderPolicy extends AbstractLoopRowTableRenderPolicy implements RenderPolicy {

    public LoopFullTableIncludeSubRenderPolicy() {
        this(false);
    }

    public LoopFullTableIncludeSubRenderPolicy(boolean onSameLine) {
        this("[", "]", onSameLine);
    }

    public LoopFullTableIncludeSubRenderPolicy(String prefix, String suffix) {
        this(prefix, suffix, false);
    }

    public LoopFullTableIncludeSubRenderPolicy(String prefix, String suffix, boolean onSameLine) {
        this.prefix = prefix;
        this.suffix = suffix;
        this.onSameLine = onSameLine;
    }


    public LoopFullTableIncludeSubRenderPolicy(AbstractLoopRowTableRenderPolicy policy) {
        super(policy);
    }

    public void render(ElementTemplate eleTemplate, Object data, XWPFTemplate template) {
        try {
            XWPFTableCell tagCell = this.dealPlaceTag(eleTemplate);
            int headerNumber = WordTableUtils.findCellVMergeNumber(tagCell);
            int templateRowIndex = this.getTemplateRowIndex(tagCell) + headerNumber - 1;
            XWPFTable table = tagCell.getTableRow().getTable();

            if (!(data instanceof Iterable)) {
                table.removeRow(templateRowIndex);
                return;
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

            int perPageNumber = pageLine / templateRowNumber;
            Iterator<?> iterator = ((Iterable<?>) data).iterator();
            boolean hasNext = iterator.hasNext();
            while (hasNext) {
                Object root = iterator.next();
                hasNext = iterator.hasNext();

                if (root instanceof Map) {
                    Map<?, ?> temp = (Map<?, ?>) root;
                    ((Map<?, ?>) root).forEach((k, v) -> {
                        if (k instanceof String) {
                            globalEnv.put((String) k, v);
                        }
                    });
                    if (temp.containsKey("subs")) {
                        Object o = temp.get("subs");
                        if (o instanceof Collection) {
                            int dataCount = ((Collection<?>) o).size();
                            int index = 0;
                            int tempTemplateRowIndex = 0;
                            int insertPosition;
                            int tableCount = dataCount / perPageNumber + (dataCount % perPageNumber > 0 ? 1 : 0);
                            int currentPage = 1;
                            boolean firstFlag = true;

                            Iterator<?> subIterator = ((Collection<?>) o).iterator();
                            boolean hasSubNext = subIterator.hasNext();
                            XWPFTable currentTable = table;
                            XWPFParagraph paragraph;
                            while (hasSubNext) {
                                Object sub = subIterator.next();
                                hasSubNext = subIterator.hasNext();

                                // Determine whether to cross page and copy a new table across pages
                                if (index % perPageNumber == 0) {
                                    if (index != 0) {
                                        this.removeMultipleLine(templateRowNumber, currentTable, tempTemplateRowIndex);
                                        if (isRemoveNextLine && currentTable.getRows().size() > tempTemplateRowIndex) {
                                            currentTable.removeRow(tempTemplateRowIndex);
                                        }
                                        this.renderMultipleRow(currentTable, tempTemplateRowIndex, -1, resolver, documentProcessor);
                                    }
                                    if (currentPage <= tableCount) {
                                        // set page break
                                        XmlCursor xmlCursor = table.getCTTbl().newCursor();
                                        paragraph = xwpfDocument.insertNewParagraph(xmlCursor);
                                        WordTableUtils.setPageBreak(paragraph, 1);
                                        WordTableUtils.setMinHeightParagraph(paragraph);

                                        xmlCursor.toParent();
                                        currentTable = WordTableUtils.copyToXmlCursorBefore(xwpfDocument, table, xmlCursor);
                                        xmlCursor.close();
                                        tempTemplateRowIndex = templateRowIndex;
                                        currentPage++;
                                        firstFlag = true;
                                    }
                                }

                                // Insert new rows into the original table
                                EnvIterator.makeEnv(globalEnv, ++index, index < dataCount);
                                EnvModel.of(sub, globalEnv);
                                insertPosition = tempTemplateRowIndex;
                                tempTemplateRowIndex += templateRowNumber;
                                for (int i = 0; i < templateRowNumber; i++) {
                                    int currentIndex = insertPosition + i;
                                    XWPFTableRow currentRow = currentTable.getRow(currentIndex);
                                    XWPFTableRow nextRow = currentTable.insertNewTableRow(tempTemplateRowIndex + i);
                                    nextRow = WordTableUtils.copyLineContent(currentRow, nextRow, tempTemplateRowIndex + i);
                                    if (isVMerge) {
                                        if (!firstFlag) {
                                            this.setVMerge(currentRow);
                                        } else {
                                            firstFlag = false;
                                        }
                                    }
                                    this.renderMultipleRow(currentTable, currentIndex, currentIndex, resolver, documentProcessor);
                                }
                                this.removeCurrentLineData(globalEnv, sub);
                            }

                            if (isFill && dataCount > 0) {
                                int insertLine;
                                if (currentPage == 2) {
                                    insertLine = pageLine - dataCount * templateRowNumber - reduce;
                                } else if (dataCount % perPageNumber == 0) {
                                    insertLine = 0;
                                } else {
                                    insertLine = pageLine - dataCount % perPageNumber * templateRowNumber - reduce;
                                }
                                this.fillBlankRow(insertLine, currentTable, tempTemplateRowIndex);
                                this.blankDeal(currentTable, mode, tempTemplateRowIndex, insertLine);
                                tempTemplateRowIndex += insertLine;
                            }
                            this.removeMultipleLine(templateRowNumber, currentTable, tempTemplateRowIndex);
                            if (isRemoveNextLine) {
                                currentTable.removeRow(tempTemplateRowIndex);
                            }
                            this.renderMultipleRow(currentTable, tempTemplateRowIndex, -1, resolver, documentProcessor);
                            this.removeCurrentLineData(globalEnv, root);
                        }
                    }
                }
            }
            WordTableUtils.removeTable(xwpfDocument, table);
            globalEnv.putAll(original);
        } catch (Exception e) {
            throw new RenderException("HackLoopTable for " + eleTemplate + " error: " + e.getMessage(), e);
        }
    }

    protected void afterloop(XWPFTable table, Object data) {
    }
}
