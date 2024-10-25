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
import com.deepoove.poi.xwpf.NiceXWPFDocument;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTVMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;

import java.util.List;
import java.util.Map;

public class LoopFullTableInsertFillRenderPolicy implements RenderPolicy {

    private String prefix;
    private String suffix;
    private boolean onSameLine;

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

    @Override
    public void render(ElementTemplate eleTemplate, Object data, XWPFTemplate template) {
        RunTemplate runTemplate = (RunTemplate) eleTemplate;
        XWPFRun run = runTemplate.getRun();
        try {
            if (!TableTools.isInsideTable(run)) {
                throw new IllegalStateException("The template tag " + runTemplate.getSource() + " must be inside a table");
            }
            XWPFTableCell tagCell = (XWPFTableCell) ((XWPFParagraph) run.getParent()).getBody();
            int templateRowIndex = getTemplateRowIndex(tagCell);
            XWPFTable table = tagCell.getTableRow().getTable();
            run.setText("", 0);

            if (!(data instanceof Iterable)) {
                table.removeRow(templateRowIndex);
                return;
            }
            int dataCount = 0;
            for (Object ignore : (Iterable<?>) data) {
                dataCount++;
            }

            Map<String, Object> globalEnv = template.getEnvModel().getEnv();
            int pageLine = 0;
            int reduce = 0;
            boolean isRemoveNextLine = false;
            Object n = globalEnv.get(eleTemplate.getTagName() + "_number");
            int mode = 1;
            try {
                if (n == null) {
                    // Subtract the default number of rows in the header by 1
                    pageLine = table.getRows().size() - 1;
                } else {
                    pageLine = Integer.parseInt(n.toString());
                }
                Object o = globalEnv.get(eleTemplate.getTagName() + "_mode");
                mode = o != null ? Integer.parseInt(o.toString()) : mode;
                Object r = globalEnv.get(eleTemplate.getTagName() + "_reduce");
                reduce = r != null ? Integer.parseInt(r.toString()) : reduce;
                Object rnl = globalEnv.get(eleTemplate.getTagName() + "_remove_next_line");
                isRemoveNextLine = rnl != null;
            } catch (NumberFormatException ignore) {
            }

            int[] a = new int[]{0, 0};

            TemplateResolver resolver = new TemplateResolver(template.getConfig().copy(prefix, suffix));
            int index = 0;
            XWPFTable nextTable = table;
            int tempTemplateRowIndex = 0;
            int insertPosition;
            int tableCount = dataCount / pageLine + (dataCount % pageLine > 0 ? 1 : 0);
            int currentTableIndex = 1;
            boolean firstFlag = true;
            // 删除表格后面空白的 XWPFParagraph
            NiceXWPFDocument xwpfDocument = removeEmptParagraph(template, table);
            for (Object root : (Iterable<?>) data) {

                // 判断是否跨页，跨页复制一份新表格
                if (index % pageLine == 0) {
                    if (index != 0) {
                        table.removeRow(tempTemplateRowIndex);
                        if (isRemoveNextLine) {
                            table.removeRow(tempTemplateRowIndex);
                        }
                    }
                    table = nextTable;
                    if (currentTableIndex < tableCount) {
                        WordTableUtils.setPageBreak(xwpfDocument);
                        nextTable = WordTableUtils.copyTable(template.getXWPFDocument(), table);
                        currentTableIndex++;
                    }
                    tempTemplateRowIndex = templateRowIndex;
                    firstFlag = true;
                }

                // 在原来的表上插入新的行
                insertPosition = tempTemplateRowIndex++;
                XWPFTableRow currentRow = table.getRow(insertPosition);
                if (!firstFlag) {
                    // update VMerge cells for non-first row
                    List<XWPFTableCell> tableCells = currentRow.getTableCells();
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

                XWPFTableRow nextRow = table.insertNewTableRow(tempTemplateRowIndex);
                nextRow = WordTableUtils.copyLineContent(currentRow, nextRow, tempTemplateRowIndex);
                EnvIterator.makeEnv(globalEnv, ++index, index < dataCount);
                Configure config = template.getConfig();
                config.setRenderDataComputeFactory(model -> new SpELRenderDataCompute(model, false));
                RenderDataCompute dataCompute = config.getRenderDataComputeFactory().newCompute(EnvModel.of(root, globalEnv));
                List<XWPFTableCell> cells = currentRow.getTableCells();
                cells.forEach(cell -> {
                    List<MetaTemplate> templates = resolver.resolveBodyElements(cell.getBodyElements());
                    new DocumentProcessor(template, resolver, dataCompute).process(templates);

                });
            }

            table.removeRow(tempTemplateRowIndex);
            int insertLine;
            if (dataCount <= pageLine) {
                insertLine = pageLine - dataCount;
            } else if (dataCount % pageLine == 0) {
                insertLine = 0;
            } else {
                insertLine = pageLine - dataCount % pageLine - reduce;
            }
            this.fillBlankRow(insertLine, table, tempTemplateRowIndex);
            int endRow = tempTemplateRowIndex + insertLine;
            if (isRemoveNextLine) {
                table.removeRow(tempTemplateRowIndex);
                endRow--;
            }
            // Default blank line filling, fill blank lines with a reverse slash by mode equal 2
            if (mode != 1 && insertLine != 0) {
                WordTableUtils.mergeMutipleLine(table, tempTemplateRowIndex, endRow);
                // Set diagonal border
                XWPFTableCell cellRow00 = table.getRow(tempTemplateRowIndex).getCell(0);
                WordTableUtils.setDiagonalBorder(cellRow00);
                WordTableUtils.setCellWidth(cellRow00, table.getWidth());
            }
            afterloop(table, data);

            removeEmptParagraph(template, table);
            template.reloadSelf();
        } catch (Exception e) {
            throw new RenderException("HackLoopTable for " + eleTemplate + " error: " + e.getMessage(), e);
        }
    }

    private static NiceXWPFDocument removeEmptParagraph(XWPFTemplate template, XWPFTable table) {
        NiceXWPFDocument xwpfDocument = template.getXWPFDocument();
        int posOfTable = xwpfDocument.getPosOfTable(table);
        if ((posOfTable + 1) < xwpfDocument.getBodyElements().size()) {
            IBodyElement iBodyElement = xwpfDocument.getBodyElements().get(posOfTable + 1);
            if (iBodyElement instanceof XWPFParagraph) {
                XWPFParagraph paragraph = (XWPFParagraph) iBodyElement;
                WordTableUtils.removeEmptyParagraph(paragraph);
            }
        }
        return xwpfDocument;
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
    private void fillBlankRow(int insertLine, XWPFTable table, int startIndex) {
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


}
