package com.deepoove.poi.plugin.table;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import com.deepoove.poi.exception.RenderException;
import com.deepoove.poi.policy.RenderPolicy;
import com.deepoove.poi.render.compute.EnvModel;
import com.deepoove.poi.render.compute.RenderDataCompute;
import com.deepoove.poi.render.compute.SpELRenderDataCompute;
import com.deepoove.poi.template.ElementTemplate;
import com.deepoove.poi.template.MetaTemplate;
import com.deepoove.poi.template.run.RunTemplate;
import com.deepoove.poi.util.TableTools;
import com.deepoove.poi.util.WordTableUtils;
import org.apache.poi.xwpf.usermodel.*;

import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

public class RemoveTableRowRenderPolicy implements RenderPolicy {

    private final String defaultDeleteValue;

    public RemoveTableRowRenderPolicy() {
        this(null);
    }

    public RemoveTableRowRenderPolicy(String defaultDeleteValue) {
        this.defaultDeleteValue = defaultDeleteValue;
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
            run.setText("", 0);

            Map<String, Object> globalEnv = template.getEnvModel().getEnv();
            Configure config = template.getConfig();
            config.setRenderDataComputeFactory(model -> new SpELRenderDataCompute(model, false));
            RenderDataCompute renderDataCompute = config.getRenderDataComputeFactory().newCompute(EnvModel.of(null, globalEnv));
            Object compute = renderDataCompute.compute(eleTemplate.getTagName());
            XWPFParagraph paragraph = (XWPFParagraph) run.getParent();
            XWPFTableCell tagCell = (XWPFTableCell) ((XWPFParagraph) run.getParent()).getBody();
            tagCell.removeParagraph(tagCell.getParagraphs().indexOf(paragraph));
            XWPFTableRow tableRow = tagCell.getTableRow();
            int rowIndex = WordTableUtils.findRowIndex(tableRow);
            // compute 为空 或 表达式为true 时 删除本行
            if (compute == null || compute == defaultDeleteValue) {
                removeTableCellNoSpan(tableRow, rowIndex);
            } else if (compute instanceof Boolean && Boolean.FALSE.equals(compute)) {
                removeTableCellNoSpan(tableRow, rowIndex);
            }
        } catch (Exception e) {
            throw new RenderException("Remove line failure: " + e.getMessage(), e);
        }
    }

    /**
     * 删除表格行，跨行数据需要单独处理。
     * 情况一：如果是跨行的开头则需要把数据弄到下一行对应的列，然后把他的跨行标记为 restart
     * 情况二： 如果是跨行的下一列也删除
     * 如果是跨行的n行 都删除则不应该移动restart的数据，而是删除数据
     *
     * @param tableRow 表格行
     */
    private void removeTableCellNoSpan(XWPFTableRow tableRow, int rowIndex) {
        int size = tableRow.getTableCells().size();
        XWPFTable table = tableRow.getTable();
        // 判断是否有某一列跨行，如果有则合并单元格到下一行
        boolean isHasMergedVertically = false;
        for (int i = size - 1; i > 0; i--) {
            XWPFTableCell templateCell = tableRow.getCell(i);
            // 获取是否跨行
            Integer vMerge = WordTableUtils.findVMerge(templateCell);
            if (vMerge != null && vMerge == 2) {
                // 获取跨行数
                int mergedRows = WordTableUtils.findVerticalMergedRows(table, rowIndex, i);
                WordTableUtils.copyCellToNextRow(table, rowIndex, i);
                if (mergedRows == 2) {
                    // 把跨列的标记取消掉
                    table.getRow(rowIndex + 1).getTableCells().get(i).getCTTc().getTcPr().unsetVMerge();
                }
            }
        }
        tableRow.getTable().removeRow(rowIndex);
    }


    private List<MetaTemplate> removeElementTemplate(List<MetaTemplate> oldElementTemplates, List<MetaTemplate> newTemplates) {
        HashSet<String> collect = newTemplates.stream().map(MetaTemplate::variable).collect(Collectors.toCollection(HashSet::new));
        return oldElementTemplates.stream().filter(e -> !collect.contains(e.variable())).collect(Collectors.toCollection(ArrayList::new));
    }
}
