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
import com.deepoove.poi.template.run.RunTemplate;
import com.deepoove.poi.util.TableTools;
import com.deepoove.poi.util.WordTableUtils;
import com.deepoove.poi.xwpf.NiceXWPFDocument;
import org.apache.poi.xwpf.usermodel.*;
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
        RunTemplate runTemplate = (RunTemplate) eleTemplate;
        XWPFRun run = runTemplate.getRun();
        if (!TableTools.isInsideTable(run)) {
            throw new IllegalStateException("The template tag " + runTemplate.getSource() + " must be inside a table");
        }
        try {
            XWPFTableCell tagCell = (XWPFTableCell) ((XWPFParagraph) run.getParent()).getBody();
            int headerNumber = WordTableUtils.findCellVMergeNumber(tagCell);
            int templateRowIndex = this.getTemplateRowIndex(tagCell) + headerNumber - 1;
            XWPFTable table = tagCell.getTableRow().getTable();
            run.setText("", 0);

            int dataCount;
            if (data instanceof Collection) {
                dataCount = ((Collection<?>) data).size();
            } else {
                throw new RenderException("The data type is an " + data.getClass().getSimpleName() +
                    ", and the data type must be a collection");
            }

            Map<String, Object> globalEnv = template.getEnvModel().getEnv();
            Map<String, Object> original = new HashMap<>(globalEnv);
            Configure config = template.getConfig();
            RenderDataCompute dataCompute = config.getRenderDataComputeFactory()
                .newCompute(EnvModel.of(template.getEnvModel().getRoot(), globalEnv));
            TemplateResolver resolver = new TemplateResolver(template.getConfig().copy(prefix, suffix));
            DocumentProcessor documentProcessor = new DocumentProcessor(template, resolver, dataCompute);

            int templateRowNumber = 1;
            int pageLine = 0;
            int reduce = 0;
            boolean isRemoveNextLine = false;
            Object n = globalEnv.get(eleTemplate.getTagName() + "_number");
            int mode = 1;
            boolean isFill = true;
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
                // 判断是否跨页，跨页复制一份新表格
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
                        tempTemplateRowIndex = templateRowIndex;
                        currentTableIndex++;
                        firstFlag = true;
                    }
                }

                // 在原来的表上插入新的行
                insertPosition = tempTemplateRowIndex++;
                XWPFTableRow currentRow = table.getRow(insertPosition);
                if (!firstFlag) {
                    this.setVMerge(currentRow);
                } else {
                    firstFlag = false;
                }

                XWPFTableRow nextRow = table.insertNewTableRow(tempTemplateRowIndex);
                nextRow = WordTableUtils.copyLineContent(currentRow, nextRow, tempTemplateRowIndex);
                EnvIterator.makeEnv(globalEnv, ++index, index < dataCount);
                EnvModel.of(root, globalEnv);
                this.renderMultipleRow(table, insertPosition, insertPosition, resolver, documentProcessor);
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
                if (insertLine > 0) {
                    if (mode == 2) {
                        WordTableUtils.mergeMutipleLine(table, tempTemplateRowIndex, tempTemplateRowIndex + insertLine - 1);
                        // Set diagonal border
                        XWPFTableCell cellRow00 = table.getRow(tempTemplateRowIndex).getCell(0);
                        WordTableUtils.setDiagonalBorder(cellRow00);
                        WordTableUtils.setCellWidth(cellRow00, table.getWidth());
                    } else if (mode == 3) {
                        XWPFTableRow row = table.getRow(tempTemplateRowIndex);
                        XWPFTableCell cell = row.getCell((row.getTableCells().size() - 1) / 2);
                        XWPFParagraph xwpfParagraph = cell.addParagraph();
                        xwpfParagraph.createRun().setText("以下空白");
                    }
                    tempTemplateRowIndex += insertLine;
                }

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
