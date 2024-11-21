package com.deepoove.poi.plugin.table;

import com.deepoove.poi.data.RenderData;
import com.deepoove.poi.policy.RenderPolicy;
import com.deepoove.poi.util.TlBeanUtil;
import com.deepoove.poi.util.WordTableUtils;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import java.util.Map;

public abstract class AbstractLoopRowTableRenderPolicy implements RenderPolicy {

    protected String prefix;
    protected String suffix;
    protected boolean onSameLine;
    protected boolean isSaveNextLine;

    public AbstractLoopRowTableRenderPolicy(){
    }

    public AbstractLoopRowTableRenderPolicy(AbstractLoopRowTableRenderPolicy policy){
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

    protected void fillBlankRow(int insertLine, XWPFTable table, int startIndex){
        if (insertLine <= 0) {
            return;
        }
        XWPFTableRow tempRow = table.insertNewTableRow(startIndex + 1);
        tempRow = WordTableUtils.copyLineContent(table.getRow(startIndex), tempRow, startIndex + 1);
        WordTableUtils.cleanRowTextContent(tempRow);
        startIndex++;
        for (int i = 1; i < insertLine; i++) {
            tempRow = table.insertNewTableRow(startIndex + 1);
            WordTableUtils.copyLineContent(table.getRow(startIndex), tempRow, startIndex + 1);
            startIndex++;
        }
    }

    // Set the bottom border of the table to the left border style
    protected void drawBottomBorder(int currentPage, boolean isDrawBorderOfFirstPage, XWPFTable table) {
        // Set the bottom border of the table to the left border style
        if (currentPage == 2 && isDrawBorderOfFirstPage){
            WordTableUtils.setBottomBorder(table, null);
        }
        if (currentPage > 2){
            WordTableUtils.setBottomBorder(table, null);
        }
    }

    protected int countPageNumber(int dataCount, int template_row_number, int pageLine, int firstPageLine) {
        if (dataCount * template_row_number <= firstPageLine) {
            return 1;
        }
        int firstNumber = firstPageLine / template_row_number;
        int perPageNumber = pageLine / template_row_number;
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
}
