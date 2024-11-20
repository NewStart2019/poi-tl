package com.deepoove.poi.plugin.table;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import com.deepoove.poi.data.RenderData;
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
import com.deepoove.poi.util.TlBeanUtil;
import com.deepoove.poi.util.WordTableUtils;
import com.deepoove.poi.xwpf.NiceXWPFDocument;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlCursor;

import java.util.*;

public class LoopMutilpleRowRenderPolicy implements RenderPolicy {

    private String prefix;
    private String suffix;
    private boolean onSameLine;

    public LoopMutilpleRowRenderPolicy() {
        this(false);
    }

    public LoopMutilpleRowRenderPolicy(boolean onSameLine) {
        this("[", "]", onSameLine);
    }

    public LoopMutilpleRowRenderPolicy(String prefix, String suffix) {
        this(prefix, suffix, false);
    }

    public LoopMutilpleRowRenderPolicy(String prefix, String suffix, boolean onSameLine) {
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
                throw new RenderException("The template tag " + runTemplate.getSource() + " must be inside a table");
            }
            XWPFTableCell tagCell = (XWPFTableCell) ((XWPFParagraph) run.getParent()).getBody();
            int headerNumber = WordTableUtils.findCellVMergeNumber(tagCell);
            int templateRowIndex = this.getTemplateRowIndex(tagCell) + headerNumber - 1;
            XWPFTable table = tagCell.getTableRow().getTable();
            run.setText("", 0);

            int dataCount;
            if (data instanceof Collection) {
                dataCount = ((Collection<?>) data).size();
            } else {
                throw new RenderException("The data type is an " + data.getClass().getSimpleName() +
                    ", and the data type must be a collection");
            }

            Map<String, Object> globalEnv = template.getEnvModel().getEnv();
            Map<String, Object> original = new HashMap<>(globalEnv);
            int template_row_number = 1;
            int firstPageLine = 0;
            int pageLine = 0;
            int reduce = 0;
            Object n = globalEnv.get(eleTemplate.getTagName() + "_number");
            int mode = 1;
            boolean isDrawBorderOfFirstPage = false;
            try {
                if (n == null) {
                    // Subtract the default number of rows in the header by 1
                    pageLine = table.getRows().size() - 1;
                } else {
                    pageLine = Integer.parseInt(n.toString());
                }
                Object temp = globalEnv.get(eleTemplate.getTagName() + "_row_number");
                template_row_number = temp == null ? template_row_number : Integer.parseInt(temp.toString());
                temp = globalEnv.get(eleTemplate.getTagName() + "_first_number");
                firstPageLine = temp != null ? Integer.parseInt(temp.toString()) : firstPageLine;
                temp = globalEnv.get(eleTemplate.getTagName() + "_mode");
                mode = temp != null ? Integer.parseInt(temp.toString()) : mode;
                temp = globalEnv.get(eleTemplate.getTagName() + "_reduce");
                reduce = temp != null ? Integer.parseInt(temp.toString()) : reduce;
                temp = globalEnv.get(eleTemplate.getTagName() + "_fpdb");
                isDrawBorderOfFirstPage = temp != null;
            } catch (NumberFormatException ignore) {
            }
            if (template_row_number > firstPageLine) {
                throw new RenderException("Template rendering with more lines than the first page is not supported!");
            }
            if (firstPageLine % template_row_number != 0 || pageLine % template_row_number != 0) {
                throw new RenderException("The size of each page should be a multiple of the number of lines in the template for multi line rendering!");
            }

            Configure config = template.getConfig();
            RenderDataCompute dataCompute = config.getRenderDataComputeFactory()
                .newCompute(EnvModel.of(template.getEnvModel().getRoot(), globalEnv));
            TemplateResolver resolver = new TemplateResolver(template.getConfig().copy(prefix, suffix));
            DocumentProcessor documentProcessor = new DocumentProcessor(template, resolver, dataCompute);

            // Delete blank XWPFParagraph after the table
            NiceXWPFDocument xwpfDocument = template.getXWPFDocument();
            WordTableUtils.removeLastBlankParagraph(xwpfDocument);
            Iterator<?> iterator = ((Iterable<?>) data).iterator();
            boolean hasNext = iterator.hasNext();
            int index = 0;
            boolean firstPage = true;
            int allPage = countPageNumber(dataCount, template_row_number, pageLine, firstPageLine);
            int firstNumber = firstPageLine / template_row_number;
            int perPageNumber = pageLine / template_row_number;
            int currentPage = 1;
            XWPFTable nextTable = table;
            int templateRowIndex2 = templateRowIndex;
            int insertPosition;
            XWPFParagraph paragraph = null;
            while (hasNext) {
                Object root = iterator.next();
                hasNext = iterator.hasNext();

                firstPage = index < firstNumber;
                if (index == 0 || index == firstNumber || ((index > firstNumber) && (index - firstNumber) % perPageNumber == 0)) {
                    if (index != 0) {
                        removeMutipleLine(template_row_number, table, templateRowIndex);
                    }
                    // Set the bottom border of the table to the left border style
                    drawBottomBorder(currentPage, isDrawBorderOfFirstPage, table);
                    table = nextTable;
                    if (currentPage <= allPage) {
                        // set page break
                        XmlCursor xmlCursor = table.getCTTbl().newCursor();
                        xmlCursor.toNextSibling();
                        paragraph = xwpfDocument.insertNewParagraph(xmlCursor);
                        WordTableUtils.setPageBreak(paragraph, 1);
                        WordTableUtils.setMinHeightParagraph(paragraph);
                        if (firstPage) {
                            xmlCursor.toParent();
                            xmlCursor.toNextSibling();
                            nextTable = xwpfDocument.insertNewTbl(xmlCursor);
                            nextTable.removeRow(0);
                            int rowIndex = WordTableUtils.findRowIndex(tagCell);
                            templateRowIndex2 = headerNumber;
                            int temp = 0;
                            for (int i = rowIndex; i < rowIndex + headerNumber + template_row_number; i++) {
                                WordTableUtils.copyLineContent(table.getRow(i), nextTable.insertNewTableRow(temp), temp++);
                            }
                        } else {
                            xmlCursor.toParent();
                            nextTable = WordTableUtils.copyTable(xwpfDocument, table, xmlCursor);
                            templateRowIndex = templateRowIndex2;
                        }
                        xmlCursor.close();
                        currentPage++;
                    }
                }

                insertPosition = templateRowIndex;
                templateRowIndex += template_row_number;
                EnvIterator.makeEnv(globalEnv, ++index, index < dataCount);
                EnvModel.of(root, globalEnv);
                for (int i = 0; i < template_row_number; i++) {
                    XWPFTableRow currentRow = table.getRow(insertPosition + i);
                    XWPFTableRow nextRow = table.insertNewTableRow(templateRowIndex + i);
                    nextRow = WordTableUtils.copyLineContent(currentRow, nextRow, templateRowIndex + i);
                    currentRow.getTableCells().forEach(cell -> {
                        List<MetaTemplate> templates = resolver.resolveBodyElements(cell.getBodyElements());
                        documentProcessor.process(templates);
                    });
                }

                removeCurrentLineData(globalEnv, root);

            }

            if (paragraph != null) {
                WordTableUtils.removeParagraph(paragraph);
            }
            int insertLine;
            if (firstPage) {
                insertLine = firstPageLine - index * template_row_number - reduce;
            } else if ((dataCount - firstNumber) % perPageNumber == 0) {
                insertLine = 0;
            } else {
                insertLine = pageLine - (dataCount - firstNumber) % perPageNumber * template_row_number - reduce;
            }
            this.fillBlankRow(insertLine, table, templateRowIndex);

            // Default blank line filling, fill blank lines with a reverse slash by mode equal 2
            // Fill in the following blanks for Mode 3 mode fill "以下空白"
            if (insertLine > 0) {
                if (mode == 2) {
                    WordTableUtils.mergeMutipleLine(table, templateRowIndex, templateRowIndex + insertLine - 1);
                    // Set diagonal border
                    XWPFTableCell cellRow00 = table.getRow(templateRowIndex).getCell(0);
                    WordTableUtils.setDiagonalBorder(cellRow00);
                    WordTableUtils.setCellWidth(cellRow00, table.getWidth());
                } else if (mode == 3) {
                    XWPFTableRow row = table.getRow(templateRowIndex);
                    XWPFTableCell cell = row.getCell((row.getTableCells().size() - 1) / 2);
                    XWPFParagraph xwpfParagraph = cell.addParagraph();
                    xwpfParagraph.createRun().setText("以下空白");
                }
            }

            if (table != nextTable) {
                WordTableUtils.removeTable(xwpfDocument, nextTable);
            }
            removeMutipleLine(template_row_number, table, templateRowIndex + insertLine);
            afterloop(table, data);
            drawBottomBorder(currentPage, isDrawBorderOfFirstPage, table);
            globalEnv.putAll(original);
        } catch (Exception e) {
            throw new RenderException("HackLoopTable for " + eleTemplate + " error: " + e.getMessage(), e);
        }
    }

    private static void drawBottomBorder(int currentPage, boolean isDrawBorderOfFirstPage, XWPFTable table) {
        // Set the bottom border of the table to the left border style
        if (currentPage == 2 && isDrawBorderOfFirstPage){
            WordTableUtils.setBottomBorder(table, null);
        }
        if (currentPage > 2){
            WordTableUtils.setBottomBorder(table, null);
        }
    }

    private static void removeMutipleLine(int template_row_number, XWPFTable table, int templateRowIndex) {
        for (int i = templateRowIndex + template_row_number - 1; i >= templateRowIndex; i--) {
            table.removeRow(templateRowIndex);
        }
    }

    private int countPageNumber(int dataCount, int template_row_number, int pageLine, int firstPageLine) {
        if (dataCount * template_row_number <= firstPageLine) {
            return 1;
        }
        int firstNumber = firstPageLine / template_row_number;
        int perPageNumber = pageLine / template_row_number;
        return (dataCount - firstNumber) / perPageNumber + ((dataCount - firstNumber) % perPageNumber == 0 ? 0 : 1) + 1;
    }

    public static void removeCurrentLineData(Map<String, Object> globalEnv, Object root) {
        TlBeanUtil beanUtil = new TlBeanUtil();
        if (root instanceof String || TlBeanUtil.isPrimitive(root)) {
            return;
        }
        Map<String, Object> map = beanUtil.beanToMap(root, RenderData.class, 0);
        map.forEach((key, value) -> globalEnv.remove(key));
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
        XWPFTableRow tempRow = table.insertNewTableRow(startIndex);
        tempRow = WordTableUtils.copyLineContent(table.getRow(startIndex + 1), tempRow, startIndex);
        WordTableUtils.cleanRowTextContent(tempRow);
        for (int i = 1; i < insertLine; i++) {
            XWPFTableRow nextRow = table.insertNewTableRow(startIndex + 1);
            WordTableUtils.copyLineContent(tempRow, nextRow, ++startIndex);
        }
    }

}

