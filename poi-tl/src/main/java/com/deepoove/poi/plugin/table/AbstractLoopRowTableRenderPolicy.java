package com.deepoove.poi.plugin.table;

import com.deepoove.poi.data.RenderData;
import com.deepoove.poi.policy.RenderPolicy;
import com.deepoove.poi.render.processor.DocumentProcessor;
import com.deepoove.poi.resolver.TemplateResolver;
import com.deepoove.poi.template.ElementTemplate;
import com.deepoove.poi.template.MetaTemplate;
import com.deepoove.poi.template.run.RunTemplate;
import com.deepoove.poi.util.TableTools;
import com.deepoove.poi.util.TlBeanUtil;
import com.deepoove.poi.util.WordTableUtils;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTVMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;

import java.util.List;
import java.util.Map;

public abstract class AbstractLoopRowTableRenderPolicy implements RenderPolicy {

    protected String prefix;
    protected String suffix;
    protected boolean onSameLine;
    protected boolean isSaveNextLine;

    public AbstractLoopRowTableRenderPolicy() {
    }

    public AbstractLoopRowTableRenderPolicy(AbstractLoopRowTableRenderPolicy policy) {
        this.prefix = policy.getPrefix();
        this.suffix = policy.getSuffix();
        this.onSameLine = policy.isOnSameLine();
        this.isSaveNextLine = policy.isSaveNextLine();
    }

    public int getTemplateRowIndex(XWPFTableCell tagCell) {
        XWPFTableRow tagRow = tagCell.getTableRow();
        return onSameLine ? WordTableUtils.findRowIndex(tagRow) : (WordTableUtils.findRowIndex(tagRow) + 1);
    }

    public String getPrefix() {
        return prefix;
    }

    public void setPrefix(String prefix) {
        this.prefix = prefix;
    }

    public String getSuffix() {
        return suffix;
    }

    public void setSuffix(String suffix) {
        this.suffix = suffix;
    }

    public boolean isOnSameLine() {
        return onSameLine;
    }

    public void setOnSameLine(boolean onSameLine) {
        this.onSameLine = onSameLine;
    }

    public boolean isSaveNextLine() {
        return isSaveNextLine;
    }

    public void setSaveNextLine(boolean saveNextLine) {
        isSaveNextLine = saveNextLine;
    }

    // Processing placeholder symbol labels，return the template tag cell
    protected XWPFTableCell dealPlaceTag(ElementTemplate eleTemplate) {
        RunTemplate runTemplate = (RunTemplate) eleTemplate;
        XWPFRun run = runTemplate.getRun();
        if (!TableTools.isInsideTable(run)) {
            throw new IllegalStateException(
                "The template tag " + runTemplate.getSource() + " must be inside a table");
        }
        XWPFTableCell tagCell = (XWPFTableCell) ((XWPFParagraph) run.getParent()).getBody();
        run.setText("", 0);
        return tagCell;
    }

    /**
     * Insert n rows before starting the index line.
     *
     * @param insertLine insert rows
     * @param table      {@link XWPFTable table}
     * @param startIndex starting the index
     */
    protected void fillBlankRow(int insertLine, XWPFTable table, int startIndex) {
        if (insertLine <= 0 || table == null) {
            return;
        }
        XWPFTableRow tempRow = table.insertNewTableRow(startIndex);
        tempRow = WordTableUtils.copyLineContent(table.getRow(startIndex + 1), tempRow, startIndex);
        WordTableUtils.cleanRowTextContent(tempRow);
        // Remove cross row
        List<XWPFTableCell> tableCells = tempRow.getTableCells();
        for (int i = 0; i < tableCells.size(); i++) {
            WordTableUtils.unVMergeCells(tempRow, i);
        }
        for (int i = 1; i < insertLine; i++) {
            XWPFTableRow nextRow = table.insertNewTableRow(startIndex + 1);
            WordTableUtils.copyLineContent(tempRow, nextRow, ++startIndex);
        }
    }

    // Set the bottom border of the table to the left border style
    protected void drawBottomBorder(int currentPage, boolean isDrawBorderOfFirstPage, XWPFTable table) {
        // Set the bottom border of the table to the left border style
        if (currentPage == 2 && isDrawBorderOfFirstPage) {
            WordTableUtils.setBottomBorder(table, null);
        }
        if (currentPage > 2) {
            WordTableUtils.setBottomBorder(table, null);
        }
    }

    /**
     * Count the number of pages
     *
     * @param dataCount         data count
     * @param templateRowNumber template row number
     * @param pageLine          page line
     * @param firstPageLine     Number of lines that can be written on the first page
     * @return total number of pages
     */
    protected int countPageNumber(int dataCount, int templateRowNumber, int pageLine, int firstPageLine) {
        if (dataCount * templateRowNumber <= firstPageLine) {
            return 1;
        }
        int firstNumber = firstPageLine / templateRowNumber;
        int perPageNumber = pageLine / templateRowNumber;
        return (dataCount - firstNumber) / perPageNumber + ((dataCount - firstNumber) % perPageNumber == 0 ? 0 : 1) + 1;
    }

    protected void removeCurrentLineData(Map<String, Object> globalEnv, Object root) {
        TlBeanUtil beanUtil = new TlBeanUtil();
        if (root instanceof String || TlBeanUtil.isPrimitive(root)) {
            return;
        }
        Map<String, Object> map = beanUtil.beanToMap(root, RenderData.class, 0);
        map.forEach((key, value) -> globalEnv.remove(key));
    }

    /**
     * Remove n rows starting from the specified line.
     *
     * @param templateRowNumber remove rows
     * @param table             {@link XWPFTable table}
     * @param templateRowIndex  remove the starting line
     */
    protected void removeMultipleLine(int templateRowNumber, XWPFTable table, int templateRowIndex) {
        if (table == null) {
            return;
        }
        int size = table.getRows().size() - 1;
        int min = Math.min(size, templateRowIndex + templateRowNumber - 1);
        for (int i = min; i >= templateRowIndex; i--) {
            table.removeRow(templateRowIndex);
        }
    }

    /**
     * If the cross row attribute of a non first row cell is REST, set it to CONTINUE
     *
     * @param row {@link XWPFTableRow row}
     */
    protected void setVMerge(XWPFTableRow row) {
        if (row == null) {
            return;
        }
        // update VMerge cells for non-first row
        List<XWPFTableCell> tableCells = row.getTableCells();
        for (XWPFTableCell cell : tableCells) {
            CTTcPr tcPr = TableTools.getTcPr(cell);
            CTVMerge vMerge = tcPr.getVMerge();
            if (null == vMerge) continue;
            if (STMerge.RESTART == vMerge.getVal()) {
                vMerge.setVal(STMerge.CONTINUE);
            }
        }
    }

    /**
     * Rendering Multi line Template
     *
     * @param table             {@link XWPFTable table}
     * @param startIndex        starting the index
     * @param endIndex          ending the index
     * @param resolver          template  {@link TemplateResolver resolver}
     * @param documentProcessor {@link DocumentProcessor documentProcessor}
     */
    protected void renderMultipleRow(XWPFTable table, int startIndex, int endIndex, TemplateResolver resolver,
                                     DocumentProcessor documentProcessor) {
        if (endIndex < 0) {
            endIndex = table.getRows().size() + endIndex;
        }
        if (startIndex > endIndex) {
            return;
        }
        for (int i = startIndex; i <= endIndex; i++) {
            List<XWPFTableCell> cells = table.getRow(i).getTableCells();
            cells.forEach(cell -> {
                List<MetaTemplate> templates = resolver.resolveBodyElements(cell.getBodyElements());
                documentProcessor.process(templates);
            });
        }
    }

    /**
     * Blank line processing
     *
     * @param table         {@link XWPFTable table}
     * @param mode          mode， 1:blank line, 2:diagonal line, 3:text "以下空白"
     * @param startRowIndex template row index
     * @param mergeLines    merge rows
     */
    protected void blankDeal(XWPFTable table, int mode, int startRowIndex, int mergeLines) {
        blankDeal(table, mode, startRowIndex, mergeLines, true);
    }

    protected void blankDeal(XWPFTable table, int mode, int startRowIndex, int mergeLines, boolean isWriteBlank) {
        if (table == null || startRowIndex < 0 || mergeLines <= 0) {
            return;
        }
        int endIndex = startRowIndex + mergeLines - 1;
        endIndex = Math.min(endIndex, table.getRows().size() - 1);
        if (mode == 2) {
            WordTableUtils.mergeMutipleLine(table, startRowIndex, endIndex);
            // Set diagonal border
            XWPFTableCell cellRow00 = table.getRow(startRowIndex).getCell(0);
            WordTableUtils.setDiagonalBorder(cellRow00);
            WordTableUtils.setCellWidth(cellRow00, table.getWidth());
        } else if (mode == 3 && isWriteBlank) {
            XWPFTableRow row = table.getRow(startRowIndex);
            WordTableUtils.cleanRowTextContent(row);
            XWPFTableCell cell = row.getCell((row.getTableCells().size() - 1) / 2);
            XWPFParagraph xwpfParagraph = cell.addParagraph();
            xwpfParagraph.createRun().setText("以下空白");
        }
    }
}
