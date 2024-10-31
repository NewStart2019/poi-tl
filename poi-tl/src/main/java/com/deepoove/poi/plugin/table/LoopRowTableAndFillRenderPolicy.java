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

public class LoopRowTableAndFillRenderPolicy implements RenderPolicy {

    private String prefix;
    private String suffix;
    private boolean onSameLine;
    private boolean isSaveNextLine;

    public LoopRowTableAndFillRenderPolicy() {
        this(false);
    }

    public LoopRowTableAndFillRenderPolicy(boolean onSameLine) {
        this("[", "]", onSameLine);
    }

    public LoopRowTableAndFillRenderPolicy(boolean onSameLine, boolean isSaveNextLine) {
        this.prefix = "[";
        this.suffix = "]";
        this.onSameLine = onSameLine;
        this.isSaveNextLine = isSaveNextLine;
    }

    public LoopRowTableAndFillRenderPolicy(String prefix, String suffix) {
        this(prefix, suffix, false);
    }

    public LoopRowTableAndFillRenderPolicy(String prefix, String suffix, boolean onSameLine) {
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
                throw new IllegalStateException("The template tag " + runTemplate.getSource() + " must be inside a table");
            }
            XWPFTableCell tagCell = (XWPFTableCell) ((XWPFParagraph) run.getParent()).getBody();
            XWPFTable table = tagCell.getTableRow().getTable();
            run.setText("", 0);

            int oldRowNumber = table.getRows().size();

            int headerNumber = WordTableUtils.findCellVMergeNumber(tagCell);
            int templateRowIndex = getTemplateRowIndex(tagCell) + headerNumber - 1;
            Map<String, Object> globalEnv = template.getEnvModel().getEnv();
            // number of lines
            int index = 0;
            if (data instanceof Iterable) {
                Iterator<?> iterator = ((Iterable<?>) data).iterator();
                XWPFTableRow templateRow = table.getRow(templateRowIndex);
                int insertPosition = templateRowIndex;

                TemplateResolver resolver = new TemplateResolver(template.getConfig().copy(prefix, suffix));
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
                    RenderDataCompute dataCompute = config.getRenderDataComputeFactory().newCompute(EnvModel.of(root, globalEnv));
                    List<XWPFTableCell> cells = nextRow.getTableCells();
                    cells.forEach(cell -> {
                        List<MetaTemplate> templates = resolver.resolveBodyElements(cell.getBodyElements());
                        new DocumentProcessor(template, resolver, dataCompute).process(templates);
                    });
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
                Object r = globalEnv.get(eleTemplate.getTagName() + "_reduce");
                reduce = r != null ? Integer.parseInt(r.toString()) : reduce;
                Object h = globalEnv.get(eleTemplate.getTagName() + "_header");
                tableHeaderLine = h != null ? Integer.parseInt(h.toString()) : tableHeaderLine;
                Object f = globalEnv.get(eleTemplate.getTagName() + "_footer");
                tableFooterLine = f != null ? Integer.parseInt(f.toString()) : tableFooterLine;
                Object o = globalEnv.get(eleTemplate.getTagName() + "_mode");
                mode = o != null ? Integer.parseInt(o.toString()) : mode;
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
                    remain = (index - pageLine + tableHeaderLine) % pageLine;
                    insertLine = pageLine - remain;
                    if (insertLine > tableFooterLine) {
                        insertLine = insertLine - tableFooterLine - reduce;
                        this.fillBlankRow(insertLine, table, templateRowIndex, mode);
                    }
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
