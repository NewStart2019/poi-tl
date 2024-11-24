package com.deepoove.poi.plugin.table;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import com.deepoove.poi.exception.RenderException;
import com.deepoove.poi.policy.RenderPolicy;
import com.deepoove.poi.render.compute.EnvModel;
import com.deepoove.poi.render.compute.RenderDataCompute;
import com.deepoove.poi.render.processor.DocumentProcessor;
import com.deepoove.poi.render.processor.EnvIterator;
import com.deepoove.poi.resolver.TemplateResolver;
import com.deepoove.poi.template.ElementTemplate;
import com.deepoove.poi.util.WordTableUtils;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

public class LoopRowTableAndFillRenderPolicy extends AbstractLoopRowTableRenderPolicy implements RenderPolicy {

    public LoopRowTableAndFillRenderPolicy() {
        this(false);
    }

    public LoopRowTableAndFillRenderPolicy(boolean onSameLine) {
        this("[", "]", onSameLine);
    }

    public LoopRowTableAndFillRenderPolicy(boolean onSameLine, boolean isSaveNextLine) {
        super();
        this.prefix = "[";
        this.suffix = "]";
        this.onSameLine = onSameLine;
        this.isSaveNextLine = isSaveNextLine;
    }

    public LoopRowTableAndFillRenderPolicy(String prefix, String suffix) {
        this(prefix, suffix, false);
    }

    public LoopRowTableAndFillRenderPolicy(String prefix, String suffix, boolean onSameLine) {
        super();
        this.prefix = prefix;
        this.suffix = suffix;
        this.onSameLine = onSameLine;
    }

    public LoopRowTableAndFillRenderPolicy(AbstractLoopRowTableRenderPolicy policy) {
        super(policy);
    }

    @Override
    public void render(ElementTemplate eleTemplate, Object data, XWPFTemplate template) {
        try {
            XWPFTableCell tagCell = this.dealPlaceTag(eleTemplate);
            XWPFTable table = tagCell.getTableRow().getTable();

            int oldRowNumber = table.getRows().size();
            int headerNumber = WordTableUtils.findCellVMergeNumber(tagCell);
            int templateRowIndex = getTemplateRowIndex(tagCell) + headerNumber - 1;
            Map<String, Object> globalEnv = template.getEnvModel().getEnv();
            Map<String, Object> original = new HashMap<>(globalEnv);

            // number of lines
            int index = 0;
            if (data instanceof Iterable) {
                Iterator<?> iterator = ((Iterable<?>) data).iterator();
                int insertPosition;

                this.initDeal(template, globalEnv);
                boolean firstFlag = true;
                boolean hasNext = iterator.hasNext();
                while (hasNext) {
                    Object root = iterator.next();
                    hasNext = iterator.hasNext();

                    insertPosition = templateRowIndex++;
                    XWPFTableRow nextRow = table.insertNewTableRow(insertPosition);
                    nextRow = WordTableUtils.copyLineContent(table.getRow(templateRowIndex), nextRow, insertPosition);
                    if (!firstFlag) {
                        this.setVMerge(nextRow);
                    } else {
                        firstFlag = false;
                    }
                    WordTableUtils.setTableRow(table, nextRow, insertPosition);

                    EnvIterator.makeEnv(globalEnv, ++index, hasNext);
                    EnvModel.of(root, globalEnv);
                    this.renderMultipleRow(table, insertPosition, insertPosition, resolver, documentProcessor);
                    this.removeCurrentLineData(globalEnv, root);
                }
            }

            // Default template to fill a full page of the table
            int pageLine = oldRowNumber - headerNumber;
            // Fill to reduce the number of rows
            int reduce = 0;
            int tableHeaderLine = headerNumber;
            int tableFooterLine = 0;
            Object temp = globalEnv.get(eleTemplate.getTagName() + "_number");
            int mode = 1;
            boolean isFill = true;
            try {
                pageLine = temp == null ? pageLine : Integer.parseInt(temp.toString());
                temp = globalEnv.get(eleTemplate.getTagName() + "_reduce");
                reduce = temp != null ? Integer.parseInt(temp.toString()) : reduce;
                temp = globalEnv.get(eleTemplate.getTagName() + "_header");
                tableHeaderLine = temp != null ? Integer.parseInt(temp.toString()) : tableHeaderLine;
                temp = globalEnv.get(eleTemplate.getTagName() + "_footer");
                tableFooterLine = temp != null ? Integer.parseInt(temp.toString()) : tableFooterLine;
                temp = globalEnv.get(eleTemplate.getTagName() + "_mode");
                mode = temp != null ? Integer.parseInt(temp.toString()) : mode;
                temp = globalEnv.get(eleTemplate.getTagName() + "_nofill");
                isFill = temp == null;
            } catch (NumberFormatException ignore) {
            }
            if (isFill) {
                // table.removeRow(templateRowIndex);
                // The first page is sufficient to write data and the end content
                int insertLine;
                if (index < pageLine) {
                    insertLine = pageLine - index - reduce;
                    this.fillBlankRow(insertLine, table, templateRowIndex);
                    this.blankDeal(table, mode, templateRowIndex, insertLine);
                    templateRowIndex += insertLine;
                } else if (index == pageLine) {
                } else if (index < pageLine + tableFooterLine) {
                    // The first part fill bank row
                    insertLine = pageLine + tableFooterLine - index;
                    this.fillBlankRow(insertLine, table, templateRowIndex);
                    this.blankDeal(table, mode, templateRowIndex, insertLine);
                    templateRowIndex += insertLine;
                    // The second part fill blank row
                    insertLine = tableHeaderLine + pageLine - reduce;
                    this.fillBlankRow(insertLine, table, templateRowIndex);
                    this.blankDeal(table, mode, templateRowIndex, insertLine, false);
                    templateRowIndex += insertLine;
                } else if (index == (pageLine + tableFooterLine)) {
                    insertLine = tableHeaderLine + pageLine - reduce;
                    this.fillBlankRow(insertLine, table, templateRowIndex);
                    this.blankDeal(table, mode, templateRowIndex, insertLine);
                    templateRowIndex += insertLine;
                } else {
                    // Other pages
                    int perPageNumber = tableHeaderLine + pageLine + tableFooterLine;
                    int remainData = index - pageLine - tableFooterLine;
                    int remain = perPageNumber - remainData % perPageNumber;
                    if (remain > tableFooterLine) {
                        insertLine = remain - tableFooterLine - reduce;
                        this.fillBlankRow(insertLine, table, templateRowIndex);
                        this.blankDeal(table, mode, templateRowIndex, insertLine);
                        templateRowIndex += insertLine;
                    } else if (remain == tableFooterLine) {
                    } else {
                        // The first part fill bank row
                        insertLine = remain;
                        if (insertLine > 0) {
                            this.fillBlankRow(insertLine, table, templateRowIndex);
                            this.blankDeal(table, mode, templateRowIndex, insertLine);
                            templateRowIndex += insertLine;
                            insertLine = pageLine + tableHeaderLine - reduce;
                            this.fillBlankRow(insertLine, table, templateRowIndex);
                            this.blankDeal(table, mode, templateRowIndex, insertLine, false);
                            templateRowIndex += insertLine;
                        }
                    }
                }
            }
            table.removeRow(templateRowIndex);
            globalEnv.putAll(original);
            afterloop(table, data);
        } catch (Exception e) {
            throw new RenderException("HackLoopTable for " + eleTemplate + " error: " + e.getMessage(), e);
        }
    }

    protected void afterloop(XWPFTable table, Object data) {
    }
}
