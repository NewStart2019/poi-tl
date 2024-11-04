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
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlObject;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTVMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;

import java.util.Iterator;
import java.util.List;
import java.util.Map;

/**
 * Hack for loop table row
 *
 * @author Sayi
 */
public class LoopRowTableRenderPolicy implements RenderPolicy {

    private String prefix;
    private String suffix;
    private boolean onSameLine;
    private boolean isSaveNextLine;

    public LoopRowTableRenderPolicy() {
        this(false);
    }

    public LoopRowTableRenderPolicy(boolean onSameLine) {
        this("[", "]", onSameLine);
    }

    public LoopRowTableRenderPolicy(boolean onSameLine, boolean isSaveNextLine) {
        this.prefix = "[";
        this.suffix = "]";
        this.onSameLine = onSameLine;
        this.isSaveNextLine = isSaveNextLine;
    }

    public LoopRowTableRenderPolicy(String prefix, String suffix) {
        this(prefix, suffix, false);
    }

    public LoopRowTableRenderPolicy(String prefix, String suffix, boolean onSameLine) {
        this.prefix = prefix;
        this.suffix = suffix;
        this.onSameLine = onSameLine;
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

            int oldRowNumber = table.getRows().size();

            int templateRowIndex = getTemplateRowIndex(tagCell);
            Map<String, Object> globalEnv = template.getEnvModel().getEnv();
            if (data instanceof Iterable) {
                Iterator<?> iterator = ((Iterable<?>) data).iterator();
                XWPFTableRow templateRow = table.getRow(templateRowIndex);
                int insertPosition = templateRowIndex;

                TemplateResolver resolver = new TemplateResolver(template.getConfig().copy(prefix, suffix));
                boolean firstFlag = true;
                int index = 0;
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
                        // update VMerge cells for non-first row
                        List<XWPFTableCell> tableCells = nextRow.getTableCells();
                        for (XWPFTableCell cell : tableCells) {
                            CTTcPr tcPr = TableTools.getTcPr(cell);
                            CTVMerge vMerge = tcPr.getVMerge();
                            if (null == vMerge) continue;
                            if (STMerge.RESTART == vMerge.getVal()) {
                                vMerge.setVal(STMerge.CONTINUE);
                            }
                        }
                    } else {
                        firstFlag = false;
                    }
                    WordTableUtils.setTableRow(table, nextRow, insertPosition);

                    EnvIterator.makeEnv(globalEnv, ++index, hasNext);
                    Configure config = template.getConfig();
                    config.setRenderDataComputeFactory(model -> new SpELRenderDataCompute(model, false));
                    RenderDataCompute dataCompute = config.getRenderDataComputeFactory()
                        .newCompute(EnvModel.of(root, globalEnv));
                    List<XWPFTableCell> cells = nextRow.getTableCells();
                    cells.forEach(cell -> {
                        List<MetaTemplate> templates = resolver.resolveBodyElements(cell.getBodyElements());
                        new DocumentProcessor(template, resolver, dataCompute).process(templates);
                    });

                    LoopCopyHeaderRowRenderPolicy.removeCurrentLineData(globalEnv, root);
                }
            }

            table.removeRow(templateRowIndex);
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
