package com.deepoove.poi.plugin.table;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import com.deepoove.poi.exception.RenderException;
import com.deepoove.poi.policy.RenderPolicy;
import com.deepoove.poi.resolver.TemplateResolver;
import com.deepoove.poi.template.ElementTemplate;
import com.deepoove.poi.template.MetaTemplate;
import com.deepoove.poi.template.run.RunTemplate;
import com.deepoove.poi.util.TableTools;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;

import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.stream.Collectors;

public class RemoveTableRowRenderPolicy implements RenderPolicy {

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
            XWPFTableRow tableRow = tagCell.getTableRow();
            XWPFTable table = tableRow.getTable();
            int rowIndex = table.getRows().indexOf(tableRow);
            if (data instanceof Boolean) {
                Boolean d = (Boolean) data;
                if (d) {
                    // 获取这一行的所有的 标签 停止渲染
                    List<MetaTemplate> elementTemplates = template.getElementTemplates();
                    List<XWPFTableCell> tableCells = tableRow.getTableCells();
                    Configure config = template.getConfig();
                    TemplateResolver resolver = new TemplateResolver(template.getConfig().copy(config.getGramerPrefix(), config.getGramerSuffix()));
                    List<MetaTemplate> newTemplates = new ArrayList<>();
                    for (XWPFTableCell tableCell : tableCells) {
                        List<MetaTemplate> templates = resolver.resolveBodyElements(tableCell.getBodyElements());
                        newTemplates.addAll(templates);
                    }
                    template.setElementTemplates(removeElementTemplate(elementTemplates, newTemplates));
                    for (int i = 0; i < tableRow.getTableCells().size(); i++) {
                        XWPFTableCell templateCell = tableRow.getCell(i);
                        // 获取是否跨行
                        CTTcPr tcPr = templateCell.getCTTc().getTcPr();
                        if (null != tcPr && null != tcPr.getGridSpan() && tcPr.getGridSpan().getVal().intValue() == 1) {
                            tableRow.removeCell(i + 1);
                        }
                    }
                }
            }
        } catch (Exception e) {
            throw new RenderException("Remove line failure: " + e.getMessage(), e);
        }
    }

    private List<MetaTemplate> removeElementTemplate(List<MetaTemplate> oldElementTemplates, List<MetaTemplate> newTemplates) {
        HashSet<String> collect = newTemplates.stream().map(MetaTemplate::variable).collect(Collectors.toCollection(HashSet::new));
        return oldElementTemplates.stream().filter(e -> !collect.contains(e.variable())).collect(Collectors.toCollection(ArrayList::new));
    }

    private int getRowIndex(XWPFTableRow row) {
        List<XWPFTableRow> rows = row.getTable().getRows();
        return rows.indexOf(row);
    }

}
