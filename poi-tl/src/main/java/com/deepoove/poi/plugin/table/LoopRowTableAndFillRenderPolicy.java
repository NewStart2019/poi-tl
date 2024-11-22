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
import com.deepoove.poi.template.MetaTemplate;
import com.deepoove.poi.template.run.RunTemplate;
import com.deepoove.poi.util.TableTools;
import com.deepoove.poi.util.WordTableUtils;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlObject;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTVMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;

import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
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
                XWPFTableRow templateRow = table.getRow(templateRowIndex);
                int insertPosition = templateRowIndex;

                Configure config = template.getConfig();
                RenderDataCompute dataCompute = config.getRenderDataComputeFactory()
                    .newCompute(EnvModel.of(template.getEnvModel().getRoot(), globalEnv));
                TemplateResolver resolver = new TemplateResolver(template.getConfig().copy(prefix, suffix));
                DocumentProcessor documentProcessor = new DocumentProcessor(template, resolver, dataCompute);

                boolean firstFlag = true;
                boolean hasNext = iterator.hasNext();
                while (hasNext) {
                    Object root = iterator.next();
                    hasNext = iterator.hasNext();

                    insertPosition = templateRowIndex++;
                    XWPFTableRow nextRow = table.insertNewTableRow(insertPosition);
                    WordTableUtils.setTableRow(table, templateRow, insertPosition);

                    // double set row
                    XmlCursor newCursor = templateRow.getCtRow().newCursor();
                    newCursor.toPrevSibling();
                    XmlObject object = newCursor.getObject();
                    nextRow = new XWPFTableRow((CTRow) object, table);
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
            int pageLine = oldRowNumber;
            // Fill to reduce the number of rows
            int reduce = 0;
            int tableHeaderLine = 0;
            int tableFooterLine = 0;
            Object n = globalEnv.get(eleTemplate.getTagName() + "_number");
            boolean isSaveNextLine = true;
            int mode = 1;
            try {
                pageLine = n == null ? pageLine : Integer.parseInt(n.toString());
                n = globalEnv.get(eleTemplate.getTagName() + "_reduce");
                reduce = n != null ? Integer.parseInt(n.toString()) : reduce;
                n = globalEnv.get(eleTemplate.getTagName() + "_header");
                tableHeaderLine = n != null ? Integer.parseInt(n.toString()) : tableHeaderLine;
                n = globalEnv.get(eleTemplate.getTagName() + "_footer");
                tableFooterLine = n != null ? Integer.parseInt(n.toString()) : tableFooterLine;
                n = globalEnv.get(eleTemplate.getTagName() + "_mode");
                mode = n != null ? Integer.parseInt(n.toString()) : mode;
            } catch (NumberFormatException ignore) {
            }
            if (n != null) {
                table.removeRow(templateRowIndex);
                // Do not save the next line
                if (!isSaveNextLine) {
                    templateRowIndex -= 1;
                }
                // Get the number of lines that can be written on the first page
                int firstPageLine = pageLine - tableHeaderLine - tableFooterLine;
                int remain = 0;
                int insertLine = 0;
                // no cross page
                if (firstPageLine >= index) {
                    remain = index;
                    insertLine = firstPageLine - remain - reduce;
                    this.fillBlankRow(insertLine, table, templateRowIndex, mode);
                } else {
                    // 第一页可写行数
                    firstPageLine = pageLine - tableHeaderLine;
                    // 判断超过第一页
                    if (index + tableFooterLine > firstPageLine) {
                        // 除了第一页剩余的行数
                        int remain1 = index - firstPageLine;
                        // 最后一页可写行数
                        if (remain1 % pageLine == 0) {
                            insertLine = pageLine - tableFooterLine - reduce;
                        } else {
                            int temp = remain1 % pageLine;
                            // 剩余行数超过一页
                            if (temp + tableFooterLine > pageLine) {
                                insertLine = pageLine - remain + (pageLine - tableFooterLine) - reduce;
                            } else {
                                insertLine = pageLine - temp - tableFooterLine - reduce;
                            }
                        }
                    }
                    this.fillBlankRow(insertLine, table, templateRowIndex, mode);
                }
                // Fill blank lines with a reverse slash
                if (mode != 1) {
                    WordTableUtils.mergeMutipleLine(table, templateRowIndex, templateRowIndex + insertLine);
                    // Set diagonal border
                    XWPFTableCell cellRow00 = table.getRow(templateRowIndex).getCell(0);
                    WordTableUtils.setDiagonalBorder(cellRow00);
                    WordTableUtils.setCellWidth(cellRow00, table.getWidth());
                }
            } else {
                table.removeRow(templateRowIndex);
            }

            globalEnv.putAll(original);
            afterloop(table, data);
        } catch (Exception e) {
            throw new RenderException("HackLoopTable for " + eleTemplate + " error: " + e.getMessage(), e);
        }
    }

    protected void afterloop(XWPFTable table, Object data) {
    }

    /**
     * Fill the blank row
     *
     * @param insertLine The number of rows per page
     * @param table      XWPFTable
     * @param startIndex Start writing the position of blank lines
     */
    private void fillBlankRow(int insertLine, XWPFTable table, int startIndex, int mode) {
        if (insertLine <= 0) {
            return;
        }
        if (mode != 1) {
            // Mode 2 requires splitting across rows and merging cells before reducing the width
            XWPFTableRow row = table.getRow(startIndex);
            int size = row.getTableCells().size();
            for (int i = 0; i < size; i++) {
                WordTableUtils.unVMergeCells(row, i);
            }
        }
        startIndex += 1;
        XWPFTableRow tempRow = table.insertNewTableRow(startIndex);
        tempRow = WordTableUtils.copyLineContent(table.getRow(startIndex - 1), tempRow, startIndex);
        WordTableUtils.cleanRowTextContent(tempRow);
        for (int i = 1; i < insertLine; i++) {
            tempRow = table.insertNewTableRow(startIndex + 1);
            WordTableUtils.copyLineContent(table.getRow(startIndex), tempRow, startIndex + 1);
            startIndex += 1;
        }
    }


}
