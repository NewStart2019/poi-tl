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
import com.deepoove.poi.render.compute.RenderDataCompute;
import com.deepoove.poi.render.processor.DocumentProcessor;
import com.deepoove.poi.render.processor.EnvIterator;
import com.deepoove.poi.resolver.TemplateResolver;
import com.deepoove.poi.template.ElementTemplate;
import com.deepoove.poi.template.MetaTemplate;
import com.deepoove.poi.template.run.RunTemplate;
import com.deepoove.poi.util.ReflectionUtils;
import com.deepoove.poi.util.TableTools;
import com.deepoove.poi.util.WordTableUtils;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlObject;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;

import java.util.Collections;
import java.util.Iterator;
import java.util.List;

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
                        this.copyLine(templateRow, table.getRow(templateRowIndex + 1), templateRowIndex + 1);
                    }
                    this.copyLine(currentLine, templateRow, templateRowIndex);

                    RenderDataCompute dataCompute = template.getConfig()
                        .getRenderDataComputeFactory()
                        .newCompute(EnvModel.of(root, EnvIterator.makeEnv(index++, hasNext)));
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
                        this.cleanRowTextContent(templateRow);
                        this.copyLine(row, templateRow, templateRowIndex);
                        this.cleanRowTextContent(row);
                    } else if (newAdd == 1) {
                        XWPFTableRow row = table.getRow(templateRowIndex + 1);
                        this.cleanRowTextContent(templateRow);
                        this.copyLine(row, templateRow, templateRowIndex);
                        table.removeRow(templateRowIndex + 1);
                    } else {
                        table.removeRow(templateRowIndex + 1);
                        table.removeRow(templateRowIndex);
                    }
                } else {
                    if (newAdd == 0) {
                        this.cleanRowTextContent(templateRow);
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

    /**
     * Copy the content of the current line to the next line. If the next line is a newly added line,
     * then directly copy the entire XML of the current line to the next line. Otherwise, just copy
     * the content to the next line.
     *
     * @param currentLine      current line
     * @param nextLine         next line
     * @param templateRowIndex next line row index
     */
    private void copyLine(XWPFTableRow currentLine, XWPFTableRow nextLine, int templateRowIndex) {
        XWPFTable table = currentLine.getTable();
        if (CollectionUtils.isEmpty(nextLine.getTableCells())) {
            // 复制行
            XmlCursor sourceCursor = currentLine.getCtRow().newCursor();
            XmlObject object = sourceCursor.getObject();
            XmlObject targetXmlObject = nextLine.getCtRow().newCursor().getObject();
            targetXmlObject = targetXmlObject.set(object);
            nextLine = new XWPFTableRow((CTRow) targetXmlObject, table);
            setTableRow(table, nextLine, templateRowIndex);
        } else {
            List<XWPFTableCell> tableCells = currentLine.getTableCells();
            int nextCellSize = nextLine.getTableCells().size() - 1;
            for (int i = 0; i < tableCells.size(); i++) {
                XWPFTableCell currentCell = tableCells.get(i);
                boolean isNoHasCell = nextCellSize < i;
                XWPFTableCell nextCell = isNoHasCell ? nextLine.addNewTableCell() : nextLine.getCell(i);
                copyCellContent(currentCell, nextCell, isNoHasCell);
            }
            if (nextLine.getCtRow().isSetTrPr()) {
                nextLine.getCtRow().unsetTrPr();
            }
//            nextLine.getCtRow().setTrPr(currentLine.getCtRow().getTrPr());
        }

    }

    /**
     * 复制跨列的单元格内容包括样式，由于跨列的数据只在第一一行有，所以不需要清除目标单元格的内容
     *
     * @param source source
     * @param target target
     */
    public void copyCellContent(XWPFTableCell source, XWPFTableCell target, boolean isIncludeStyle) {
        List<XWPFParagraph> paragraphs = source.getParagraphs();
        CTPPr targetCtPPr = null;
        XWPFParagraph firstParagraph = null;
        if (CollectionUtils.isNotEmpty(target.getParagraphs())) {
            firstParagraph = target.getParagraphs().get(0);
            targetCtPPr = firstParagraph.getCTP().getPPr();
        }
        for (XWPFParagraph paragraph : source.getParagraphs()) {
            XWPFParagraph newParagraph = target.addParagraph();
            WordTableUtils.copyParagraph(paragraph, newParagraph);
            if (targetCtPPr != null) {
                // 复制段落样式
                newParagraph.getCTP().setPPr(targetCtPPr);
                newParagraph.setStyle(firstParagraph.getStyle());
            }
            if (isIncludeStyle) {
                newParagraph.getCTP().unsetPPr();
                newParagraph.getCTP().setPPr(paragraph.getCTP().getPPr());
            }
        }
        if (isIncludeStyle) {
            CTTcPr sourceTcPr = source.getCTTc().getTcPr();
            if (sourceTcPr != null) {
                target.getCTTc().setTcPr(source.getCTTc().getTcPr());
            }
        }
        if (firstParagraph != null) {
            WordTableUtils.removeParagraph(target, firstParagraph);
        }
    }

    public void cleanRowTextContent(XWPFTable table, int rowIndex) {
        cleanRowTextContent(table.getRow(rowIndex));
    }

    /**
     * 清除一行的所有文本内容
     *
     * @param templateRow {@link XWPFTableRow templateRow}
     */
    public void cleanRowTextContent(XWPFTableRow templateRow) {
        List<XWPFTableCell> tableCells = templateRow.getTableCells();
        tableCells.forEach(cell -> {
            if (CollectionUtils.isNotEmpty(cell.getParagraphs())) {
                cell.getParagraphs().forEach(WordTableUtils::removeAllRun);
            }
        });
    }

    private int getTemplateRowIndex(XWPFTableCell tagCell) {
        XWPFTableRow tagRow = tagCell.getTableRow();
        return onSameLine ? getRowIndex(tagRow) : (getRowIndex(tagRow) + 1);
    }

    protected void afterloop(XWPFTable table, Object data) {
    }

    @SuppressWarnings("unchecked")
    private void setTableRow(XWPFTable table, XWPFTableRow templateRow, int pos) {
        List<XWPFTableRow> rows = (List<XWPFTableRow>) ReflectionUtils.getValue("tableRows", table);
        rows.set(pos, templateRow);
        table.getCTTbl().setTrArray(pos, templateRow.getCtRow());
    }

    private int getRowIndex(XWPFTableRow row) {
        List<XWPFTableRow> rows = row.getTable().getRows();
        return rows.indexOf(row);
    }

}
