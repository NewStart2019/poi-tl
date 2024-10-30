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

import java.util.Collection;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

public class LoopCopyHeaderRowRenderPolicy implements RenderPolicy {

    private String prefix;
    private String suffix;
    private boolean onSameLine;

    public LoopCopyHeaderRowRenderPolicy() {
        this(false);
    }

    public LoopCopyHeaderRowRenderPolicy(boolean onSameLine) {
        this("[", "]", onSameLine);
    }

    public LoopCopyHeaderRowRenderPolicy(String prefix, String suffix) {
        this(prefix, suffix, false);
    }

    public LoopCopyHeaderRowRenderPolicy(String prefix, String suffix, boolean onSameLine) {
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

            int dataCount = 0;
            if (data instanceof Collection) {
                dataCount = ((Collection<?>) data).size();
            } else {
                return;
            }

            Map<String, Object> globalEnv = template.getEnvModel().getEnv();
            int firstPageLine = 0;
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
                Object fn = globalEnv.get(eleTemplate.getTagName() + "_first_number");
                firstPageLine = fn != null ? Integer.parseInt(fn.toString()) : 0;
                Object o = globalEnv.get(eleTemplate.getTagName() + "_mode");
                mode = o != null ? Integer.parseInt(o.toString()) : mode;
                Object r = globalEnv.get(eleTemplate.getTagName() + "_reduce");
                reduce = r != null ? Integer.parseInt(r.toString()) : reduce;
                Object rnl = globalEnv.get(eleTemplate.getTagName() + "_remove_next_line");
                isRemoveNextLine = rnl != null;
            } catch (NumberFormatException ignore) {
            }

            Configure config = template.getConfig();
            config.setRenderDataComputeFactory(model -> new SpELRenderDataCompute(model, false));
            RenderDataCompute dataCompute = null;

            TemplateResolver resolver = new TemplateResolver(template.getConfig().copy(prefix, suffix));
            // Delete blank XWPFParagraph after the table
            NiceXWPFDocument xwpfDocument = removeEmptParagraph(template, table);
            Iterator<?> iterator = ((Iterable<?>) data).iterator();
            boolean hasNext = iterator.hasNext();
            int index = 0;
            boolean firstFlag = true;
            boolean firstPage = true;
            int allPage = countPageNumber(dataCount, pageLine, firstPageLine);
            int currentPage = 1;
            XWPFTable nextTable = table;
            int templateRowIndex2 = templateRowIndex;
            int insertPosition;
            while (hasNext) {
                Object root = iterator.next();
                hasNext = iterator.hasNext();

                firstPage = index < firstPageLine;
                if (index == 0 || index == firstPageLine || (index - firstPageLine) % pageLine == 0) {
                    if (index != 0) {
                        table.removeRow(templateRowIndex);
                        if (isRemoveNextLine) {
                            table.removeRow(templateRowIndex);
                        }
                    }
                    // 存在下一页，创建表格
                    table = nextTable;
                    if (currentPage <= allPage) {
                        if (index < firstPageLine) {
                            // WordTableUtils.setPageBreak(xwpfDocument);
                            nextTable = xwpfDocument.createTable();
                            int rowIndex = WordTableUtils.findRowIndex(tagCell);
                            int headerNumber = WordTableUtils.findVerticalMergedRows(table, tagCell);
                            templateRowIndex2 = headerNumber;
                            int temp = 0;
                            for (int i = rowIndex; i <= rowIndex + headerNumber; i++) {
                                WordTableUtils.copyLineContent(table.getRow(i), nextTable.insertNewTableRow(temp), temp++);
                            }
                            WordTableUtils.removeLastRow(nextTable);
                            WordTableUtils.copyTableTblPr(table, nextTable);
                            nextTable.getCTTbl().setTblGrid(table.getCTTbl().getTblGrid());
                        } else {
                            nextTable = WordTableUtils.copyTable(xwpfDocument, table);
                            templateRowIndex = templateRowIndex2;
                        }
                        firstFlag = true;
                        currentPage++;
                    }
                }

                insertPosition = templateRowIndex++;
                if (currentPage == allPage){
                    System.out.println("currentPage");
                }
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

                XWPFTableRow nextRow = table.insertNewTableRow(templateRowIndex);
                nextRow = WordTableUtils.copyLineContent(currentRow, nextRow, templateRowIndex);
                EnvIterator.makeEnv(globalEnv, index++, index < dataCount);
                dataCompute = config.getRenderDataComputeFactory().newCompute(EnvModel.of(root, globalEnv));
                List<XWPFTableCell> cells = currentRow.getTableCells();
                RenderDataCompute finalDataCompute1 = dataCompute;
                cells.forEach(cell -> {
                    List<MetaTemplate> templates = resolver.resolveBodyElements(cell.getBodyElements());
                    new DocumentProcessor(template, resolver, finalDataCompute1).process(templates);
                });
            }

            if (firstPage) {
                table.removeRow(templateRowIndex);
                if (isRemoveNextLine) {
                    if (templateRowIndex < table.getRows().size() - 1) {
                        table.removeRow(templateRowIndex);
                        templateRowIndex--;
                    }
                }
            } else {
                table.removeRow(templateRowIndex);
                if (isRemoveNextLine) {
                    if (templateRowIndex < table.getRows().size() - 1) {
                        table.removeRow(templateRowIndex);
                    }
                }
                templateRowIndex = table.getRows().size() - 1;
            }
            int insertLine;
            if (firstPage) {
                insertLine = firstPageLine - dataCount - reduce;
            } else if ((dataCount - firstPageLine) % pageLine == 0) {
                insertLine = 0;
            } else {
                insertLine = pageLine - (dataCount - firstPageLine) % pageLine - reduce;
            }
            this.fillBlankRow(insertLine, table, templateRowIndex);


            // Default blank line filling, fill blank lines with a reverse slash by mode equal 2
            if (mode != 1 && insertLine > 0) {
                WordTableUtils.mergeMutipleLine(table, templateRowIndex + 1, templateRowIndex + insertLine);
                // Set diagonal border
                XWPFTableCell cellRow00 = table.getRow(templateRowIndex + 1).getCell(0);
                WordTableUtils.setDiagonalBorder(cellRow00);
                WordTableUtils.setCellWidth(cellRow00, table.getWidth());
            }
            afterloop(table, data);
            if (table != nextTable) {
                WordTableUtils.removeTable(xwpfDocument, nextTable);
            }

            template.reloadSelf();
        } catch (Exception e) {
            throw new RenderException("HackLoopTable for " + eleTemplate + " error: " + e.getMessage(), e);
        }
    }

    private int countPageNumber(int dataCount, int pageLine, int firstPageLine) {
        if (dataCount <= firstPageLine) {
            return 1;
        }
        return (dataCount - firstPageLine) / pageLine + (dataCount - firstPageLine) % pageLine + 1;
    }

    private static NiceXWPFDocument removeEmptParagraph(XWPFTemplate template, XWPFTable table) {
        NiceXWPFDocument xwpfDocument = template.getXWPFDocument();
        int posOfTable = xwpfDocument.getPosOfTable(table);
        if ((posOfTable + 1) < xwpfDocument.getBodyElements().size()) {
            IBodyElement iBodyElement = xwpfDocument.getBodyElements().get(posOfTable + 1);
            if (iBodyElement instanceof XWPFParagraph) {
                XWPFParagraph paragraph = (XWPFParagraph) iBodyElement;
                WordTableUtils.removeParagraph(paragraph);
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
