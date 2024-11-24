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
import com.deepoove.poi.exception.RenderException;
import com.deepoove.poi.policy.RenderPolicy;
import com.deepoove.poi.render.compute.EnvModel;
import com.deepoove.poi.render.processor.EnvIterator;
import com.deepoove.poi.template.ElementTemplate;
import com.deepoove.poi.util.WordTableUtils;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

/**
 * Render existing lines
 *
 * @author zqh
 */
public class LoopExistedRowTableRenderPolicy extends AbstractLoopRowTableRenderPolicy implements RenderPolicy {

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
        try {
            XWPFTableCell tagCell = this.dealPlaceTag(eleTemplate);
            XWPFTable table = tagCell.getTableRow().getTable();

            int headerNumber = WordTableUtils.findCellVMergeNumber(tagCell);
            int templateRowIndex = this.getTemplateRowIndex(tagCell) + headerNumber - 1;
            int allRowNumber = table.getRows().size() - headerNumber;
            XWPFTableRow templateRow;
            int index = 0;
            Map<String, Object> globalEnv = template.getEnvModel().getEnv();
            Map<String, Object> original = new HashMap<>(globalEnv);
            this.initDeal(template, globalEnv);
            // Clear the content of this template line and move the nearest line up one space
            // Default template to fill a full page of the table
            int pageLine = allRowNumber + 1;
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
            boolean firstFlag = true;
            if (data instanceof Iterable) {
                Iterator<?> iterator = ((Iterable<?>) data).iterator();
                int insertPosition;

                boolean hasNext = iterator.hasNext();
                while (hasNext) {
                    Object root = iterator.next();
                    hasNext = iterator.hasNext();
                    insertPosition = templateRowIndex++;
                    if (allRowNumber - 1 <= index) {
                        templateRow = table.insertNewTableRow(templateRowIndex);
                    } else {
                        templateRow = table.getRow(templateRowIndex);
                    }
                    XWPFTableRow currentLine = table.getRow(insertPosition);
                    templateRow = WordTableUtils.copyLineContent(currentLine, templateRow, templateRowIndex);
                    if (!firstFlag) {
                        this.setVMerge(templateRow);
                    } else {
                        firstFlag = false;
                    }

                    EnvIterator.makeEnv(globalEnv, ++index, hasNext);
                    EnvModel.of(root, globalEnv);
                    this.renderMultipleRow(table, insertPosition, insertPosition, resolver, documentProcessor);
                    this.removeCurrentLineData(globalEnv, root);
                }
            }

            if (index >= allRowNumber) {
                table.removeRow(templateRowIndex);
            } else {
                WordTableUtils.cleanRowTextContent(table, templateRowIndex);
            }
            globalEnv.clear();
            globalEnv.putAll(original);
            afterloop(table, data);
        } catch (Exception e) {
            throw new RenderException("HackLoopTable for " + eleTemplate + " error: " + e.getMessage(), e);
        }
    }


    protected void afterloop(XWPFTable table, Object data) {
    }

}
