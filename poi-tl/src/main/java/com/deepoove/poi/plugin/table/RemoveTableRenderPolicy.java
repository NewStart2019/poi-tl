package com.deepoove.poi.plugin.table;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import com.deepoove.poi.exception.RenderException;
import com.deepoove.poi.policy.RenderPolicy;
import com.deepoove.poi.render.compute.EnvModel;
import com.deepoove.poi.render.compute.RenderDataCompute;
import com.deepoove.poi.template.ElementTemplate;
import com.deepoove.poi.template.run.RunTemplate;
import com.deepoove.poi.util.TableTools;
import com.deepoove.poi.util.WordTableUtils;
import org.apache.poi.xwpf.usermodel.*;

import java.util.Map;

public class RemoveTableRenderPolicy implements RenderPolicy {

    public RemoveTableRenderPolicy() {
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
            RenderDataCompute renderDataCompute = config.getRenderDataComputeFactory().newCompute(EnvModel.of(null, globalEnv));
            Object compute = renderDataCompute.compute(eleTemplate.getTagName());
            XWPFTableCell tagCell = (XWPFTableCell) ((XWPFParagraph) run.getParent()).getBody();
            XWPFTableRow tableRow = tagCell.getTableRow();
            XWPFTable table = tableRow.getTable();
            // compute 为空 或 表达式为true 时 删除表格
            if (compute == null || (compute instanceof Boolean && Boolean.FALSE.equals(compute))) {
                WordTableUtils.removeTable(template.getXWPFDocument(), table);
            }
        } catch (Exception e) {
            throw new RenderException("Remove line failure: " + e.getMessage(), e);
        }
    }
}
