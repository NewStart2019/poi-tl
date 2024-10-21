/*
 * Copyright 2014-2024 Sayi
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package com.deepoove.poi.plugin.table;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import com.deepoove.poi.exception.RenderException;
import com.deepoove.poi.policy.RenderPolicy;
import com.deepoove.poi.render.compute.EnvModel;
import com.deepoove.poi.render.compute.RenderDataCompute;
import com.deepoove.poi.render.compute.SpELRenderDataCompute;
import com.deepoove.poi.render.processor.DocumentProcessor;
import com.deepoove.poi.render.processor.EnvIterator;
import com.deepoove.poi.resolver.TemplateResolver;
import com.deepoove.poi.template.ElementTemplate;
import com.deepoove.poi.template.MetaTemplate;
import com.deepoove.poi.template.run.RunTemplate;
import com.deepoove.poi.util.TableTools;
import com.deepoove.poi.util.WordTableUtils;
import org.apache.poi.xwpf.usermodel.*;

import java.util.Iterator;
import java.util.List;
import java.util.Map;

/**
 * loop table row
 *
 * @author Sayi
 */
public class LoopExistedRowTableRenderPolicy implements RenderPolicy {

    private String prefix;
    private String suffix;
    private boolean onSameLine;
    private boolean isSaveNextLine;

    public LoopExistedRowTableRenderPolicy() {
        this(false);
    }

    public LoopExistedRowTableRenderPolicy(boolean onSameLine) {
        this("[", "]", onSameLine, false);
    }


    public LoopExistedRowTableRenderPolicy(boolean onSameLine, boolean isSaveNextLine) {
        this("[", "]", onSameLine, isSaveNextLine);
    }

    public LoopExistedRowTableRenderPolicy(String prefix, String suffix) {
        this(prefix, suffix, false, false);
    }

    public LoopExistedRowTableRenderPolicy(String prefix, String suffix, boolean onSameLine, boolean isSaveNextLine) {
        this.prefix = prefix;
        this.suffix = suffix;
        this.onSameLine = onSameLine;
        this.isSaveNextLine = isSaveNextLine;
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

            int templateRowIndex = getTemplateRowIndex(tagCell);
            int allRowNumber = table.getRows().size() - 1;
            int oldRowNumber = allRowNumber;
            TemplateResolver resolver = new TemplateResolver(template.getConfig().copy(prefix, suffix));
            XWPFTableRow templateRow = null;
            if (data instanceof Iterable) {
                Iterator<?> iterator = ((Iterable<?>) data).iterator();
                int insertPosition;

                int index = 0;
                boolean hasNext = iterator.hasNext();
                Map<String, Object> globalEnv = template.getEnvModel().getEnv();
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
                    if (isSaveNextLine) {
                        // 把下一行移到下下一行
                        if (templateRowIndex + 1 > allRowNumber) {
                            allRowNumber += 1;
                            table.insertNewTableRow(templateRowIndex + 1);
                        }
                        WordTableUtils.copyLineContent(templateRow, table.getRow(templateRowIndex + 1), templateRowIndex + 1);
                    }
                    WordTableUtils.copyLineContent(currentLine, templateRow, templateRowIndex);

                    EnvIterator.makeEnv(globalEnv, ++index, hasNext);
                    Configure config = template.getConfig();
                    config.setRenderDataComputeFactory(model -> new SpELRenderDataCompute(model, false));
                    RenderDataCompute dataCompute = config.getRenderDataComputeFactory()
                        .newCompute(EnvModel.of(root, globalEnv));
                    List<XWPFTableCell> cells = currentLine.getTableCells();
                    cells.forEach(cell -> {
                        List<MetaTemplate> templates = resolver.resolveBodyElements(cell.getBodyElements());
                        new DocumentProcessor(template, resolver, dataCompute).process(templates);
                    });
                }
            }

            // 清空这一行模板内容内容，把最近的一行往上移动一格
            if (templateRow != null) {
                int newAdd = allRowNumber - oldRowNumber;
                if (isSaveNextLine) {
                    if (newAdd == 0) {
                        XWPFTableRow row = table.getRow(templateRowIndex + 1);
                        WordTableUtils.cleanRowTextContent(templateRow);
                        WordTableUtils.copyLineContent(row, templateRow, templateRowIndex);
                        WordTableUtils.cleanRowTextContent(row);
                    } else if (newAdd == 1) {
                        XWPFTableRow row = table.getRow(templateRowIndex + 1);
                        WordTableUtils.cleanRowTextContent(templateRow);
                        WordTableUtils.copyLineContent(row, templateRow, templateRowIndex);
                        table.removeRow(templateRowIndex + 1);
                    } else {
                        table.removeRow(templateRowIndex);
                        table.removeRow(templateRowIndex);
                    }
                } else {
                    if (newAdd == 0) {
                        WordTableUtils.cleanRowTextContent(templateRow);
                    } else {
                        table.removeRow(templateRowIndex);
                    }
                }
            }
            afterloop(table, data);
        } catch (Exception e) {
            throw new RenderException("HackLoopTable for " + eleTemplate + " error: " + e.getMessage(), e);
        }
    }

    private int getTemplateRowIndex(XWPFTableCell tagCell) {
        XWPFTableRow tagRow = tagCell.getTableRow();
        return onSameLine ? WordTableUtils.findRowIndex(tagRow) : (WordTableUtils.findRowIndex(tagRow) + 1);
    }

    protected void afterloop(XWPFTable table, Object data) {
    }

}
