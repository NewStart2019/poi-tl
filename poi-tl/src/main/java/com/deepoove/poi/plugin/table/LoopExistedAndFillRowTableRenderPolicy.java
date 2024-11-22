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

import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

/**
 * TODO 重构，支持复制表头
 * 定义每页的行数，默认读取模板中的空白行
 * 复制模板行样式
 * 自定义是否填充空白行
 * 删除渲染模式1：{@link LoopExistedRowTableRenderPolicy LoopExistedRowTableRenderPolicy}
 *
 * @author zqh
 */
public class LoopExistedAndFillRowTableRenderPolicy extends AbstractLoopRowTableRenderPolicy implements RenderPolicy {

    public LoopExistedAndFillRowTableRenderPolicy() {
        this(false);
    }

    public LoopExistedAndFillRowTableRenderPolicy(boolean onSameLine) {
        this("[", "]", onSameLine, false);
    }


    public LoopExistedAndFillRowTableRenderPolicy(boolean onSameLine, boolean isSaveNextLine) {
        this("[", "]", onSameLine, isSaveNextLine);
    }

    public LoopExistedAndFillRowTableRenderPolicy(String prefix, String suffix) {
        this(prefix, suffix, false, false);
    }

    public LoopExistedAndFillRowTableRenderPolicy(String prefix, String suffix, boolean onSameLine, boolean isSaveNextLine) {
        this.prefix = prefix;
        this.suffix = suffix;
        this.onSameLine = onSameLine;
        this.isSaveNextLine = isSaveNextLine;
    }

    public LoopExistedAndFillRowTableRenderPolicy(AbstractLoopRowTableRenderPolicy policy) {
        super(policy);
    }


    @Override
    public void render(ElementTemplate eleTemplate, Object data, XWPFTemplate template) {
        RunTemplate runTemplate = (RunTemplate) eleTemplate;
        XWPFRun run = runTemplate.getRun();
        try {
            if (!TableTools.isInsideTable(run)) {
                throw new IllegalStateException(
                    "The template tag " + runTemplate.getSource() + " must be inside a table");
            }
            XWPFTableCell tagCell = (XWPFTableCell) ((XWPFParagraph) run.getParent()).getBody();
            XWPFTable table = tagCell.getTableRow().getTable();
            run.setText("", 0);

            int headerNumber = WordTableUtils.findCellVMergeNumber(tagCell);
            int templateRowIndex = this.getTemplateRowIndex(tagCell) + headerNumber - 1;
            int allRowNumber = table.getRows().size() - 1;
            int oldRowNumber = allRowNumber;
            XWPFTableRow templateRow = null;
            int index = 0;
            Map<String, Object> globalEnv = template.getEnvModel().getEnv();
            Map<String, Object> original = new HashMap<>(globalEnv);

            TemplateResolver resolver = new TemplateResolver(template.getConfig().copy(prefix, suffix));
            Configure config = template.getConfig();
            RenderDataCompute dataCompute = config.getRenderDataComputeFactory()
                .newCompute(EnvModel.of(template.getEnvModel().getRoot(), globalEnv));
            DocumentProcessor documentProcessor = new DocumentProcessor(template, resolver, dataCompute);

            // Clear the content of this template line and move the nearest line up one space
            // Default template to fill a full page of the table
            int pageLine = oldRowNumber + 1;
            int reduce = 0;
            boolean isFill = true;
            int mode = 1;
            try {
                Object n = globalEnv.get(eleTemplate.getTagName() + "_number");
                pageLine = n == null ? pageLine : Integer.parseInt(n.toString());
                Object temp = globalEnv.get(eleTemplate.getTagName() + "_reduce");
                reduce = temp != null ? Integer.parseInt(temp.toString()) : 0;
                temp = globalEnv.get(eleTemplate.getTagName() + "_nofill");
                isFill = temp == null;
                temp = globalEnv.get(eleTemplate.getTagName() + "_mode");
                mode = temp != null ? Integer.parseInt(temp.toString()) : mode;
            } catch (NumberFormatException ignore) {
            }
            if (data instanceof Iterable) {
                Iterator<?> iterator = ((Iterable<?>) data).iterator();
                int insertPosition;

                boolean hasNext = iterator.hasNext();
                while (hasNext) {
                    Object root = iterator.next();
                    hasNext = iterator.hasNext();
                    insertPosition = templateRowIndex++;
                    if (allRowNumber < templateRowIndex) {
                        allRowNumber += 1;
                        templateRow = table.insertNewTableRow(templateRowIndex);
                    } else {
                        templateRow = table.getRow(templateRowIndex);
                    }
                    XWPFTableRow currentLine = table.getRow(insertPosition);
                    WordTableUtils.copyLineContent(currentLine, templateRow, templateRowIndex);

                    EnvIterator.makeEnv(globalEnv, ++index, hasNext);
                    EnvModel.of(root, globalEnv);
                    List<XWPFTableCell> cells = currentLine.getTableCells();
                    cells.forEach(cell -> {
                        List<MetaTemplate> templates = resolver.resolveBodyElements(cell.getBodyElements());
                        documentProcessor.process(templates);
                    });

                    this.removeCurrentLineData(globalEnv, root);
                }
            }

            int newAdd = allRowNumber - oldRowNumber;
            if (templateRow != null) {
                if (isSaveNextLine) {
                    XWPFTableRow row = table.getRow(templateRowIndex + 1);
                    WordTableUtils.cleanRowTextContent(templateRow);
                    WordTableUtils.copyLineContent(row, templateRow, templateRowIndex);
                    // Determine if there is a cross page
                    int remain = (allRowNumber + 1) % pageLine;
                    if ((allRowNumber + 1) <= pageLine) {
                        WordTableUtils.cleanRowTextContent(row);
                        this.fillBlankRow(pageLine, remain, reduce, table, templateRowIndex + 1);
                    } else if (remain == 1) {
                        table.removeRow(templateRowIndex + 1);
                    } else if (remain == 2) {
                        table.removeRow(templateRowIndex);
                        table.removeRow(templateRowIndex);
                    } else {
                        table.removeRow(templateRowIndex + 1);
                        // Fill in the remaining portion，
                        this.fillBlankRow(pageLine, remain, reduce, table, templateRowIndex);
                    }
                } else {
                    if (newAdd == 0) {
                        WordTableUtils.cleanRowTextContent(templateRow);
                    } else {
                        table.removeRow(templateRowIndex);
                        templateRowIndex -= 1;
                    }
                }
            }

            globalEnv.clear();
            globalEnv.putAll(original);
            afterloop(table, data);
        } catch (Exception e) {
            throw new RenderException("HackLoopTable for " + eleTemplate + " error: " + e.getMessage(), e);
        }
    }

    /**
     * Fill the blank row
     *
     * @param pageLine   The number of rows per page
     * @param remain     Number of rows already used
     * @param reduce     the number of rows to be reduced
     * @param table      XWPFTable
     * @param startIndex Start writing the position of blank lines
     */
    protected void fillBlankRow(int pageLine, int remain, int reduce, XWPFTable table, int startIndex) {
        if (remain == 0) {
            return;
        }
        int insertLine = pageLine - remain - reduce;
        if (insertLine > 0) {
            XWPFTableRow tempRow = table.insertNewTableRow(startIndex + 1);
            tempRow = WordTableUtils.copyLineContent(table.getRow(startIndex), tempRow, startIndex + 1);
            WordTableUtils.cleanRowTextContent(tempRow);
            startIndex += 1;
        }
        for (int i = 1; i < insertLine; i++) {
            XWPFTableRow tempRow = table.insertNewTableRow(startIndex + 1);
            WordTableUtils.copyLineContent(table.getRow(startIndex), tempRow, startIndex + 1);
            startIndex += 1;
        }
    }

    protected void afterloop(XWPFTable table, Object data) {
    }
}
