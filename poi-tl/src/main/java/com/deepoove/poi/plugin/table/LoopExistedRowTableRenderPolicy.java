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
 * loop table row
 *
 * @author Sayi
 */
public class LoopExistedRowTableRenderPolicy  extends AbstractLoopRowTableRenderPolicy implements RenderPolicy {

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

    public LoopExistedRowTableRenderPolicy(AbstractLoopRowTableRenderPolicy policy) {
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
            int templateRowIndex = getTemplateRowIndex(tagCell) + headerNumber - 1;
            int allRowNumber = table.getRows().size() - 1;
            int oldRowNumber = allRowNumber;
            TemplateResolver resolver = new TemplateResolver(template.getConfig().copy(prefix, suffix));
            XWPFTableRow templateRow = null;
            Map<String, Object> globalEnv = template.getEnvModel().getEnv();
            Map<String, Object> original = new HashMap<>(globalEnv);
            Configure config = template.getConfig();
            RenderDataCompute dataCompute = config.getRenderDataComputeFactory()
                .newCompute(EnvModel.of(template.getEnvModel().getRoot(), globalEnv));
            DocumentProcessor documentProcessor = new DocumentProcessor(template, resolver, dataCompute);
            if (data instanceof Iterable) {
                Iterator<?> iterator = ((Iterable<?>) data).iterator();
                int insertPosition;

                int index = 0;
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
                    if (isSaveNextLine) {
                        // Move the next line to the next line
                        if (templateRowIndex + 1 > allRowNumber) {
                            allRowNumber += 1;
                            table.insertNewTableRow(templateRowIndex + 1);
                        }
                        WordTableUtils.copyLineContent(templateRow, table.getRow(templateRowIndex + 1), templateRowIndex + 1);
                    }
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

            // Clear the content of this template line and move the nearest line up one space
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
            globalEnv.putAll(original);
            afterloop(table, data);
        } catch (Exception e) {
            throw new RenderException("HackLoopTable for " + eleTemplate + " error: " + e.getMessage(), e);
        }
    }

    protected void afterloop(XWPFTable table, Object data) {
    }

}
