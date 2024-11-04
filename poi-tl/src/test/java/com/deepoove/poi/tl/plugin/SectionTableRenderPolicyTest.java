package com.deepoove.poi.tl.plugin;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

import com.deepoove.poi.plugin.table.RemoveTableRowRenderPolicy;
import com.deepoove.poi.template.BlockTemplate;
import com.deepoove.poi.template.IterableTemplate;
import com.deepoove.poi.template.MetaTemplate;
import com.deepoove.poi.template.run.RunTemplate;
import org.junit.jupiter.api.Test;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import com.deepoove.poi.plugin.table.RemoveTableColumnRenderPolicy;
import com.deepoove.poi.plugin.table.SectionColumnTableRenderPolicy;

public class SectionTableRenderPolicyTest {

    String resource = "src/test/resources/template/ifcol.docx";

    @Test
    public void test() throws Exception {
        Map<String, Object> data = new HashMap<>();
        data.put("r1", 12);
        data.put("r2", 0);
        data.put("r34", 0);
        data.put("A", false);
        data.put("B", true);
        // data.put("C", true);
        // data.put("D", true);
        Configure config = Configure.builder()
                .addPlugin('-', new SectionColumnTableRenderPolicy())
                .bind("ifcol", new RemoveTableColumnRenderPolicy())
                .useSpringEL(false)
                .build();
        XWPFTemplate template = XWPFTemplate.compile(resource, config).render(data);
        template.writeToFile("target/out_render_ifcol.docx");
    }
    // 初始化默认字段
    private void initFieldByDocument(List<MetaTemplate> elementTemplates, Map<String, Object> resultMap, String defaultPlaceholder) {
        elementTemplates.stream().parallel().forEach(ele -> {
            // 图片、图表等渲染不设置默认值
            if (ele instanceof RunTemplate) {
                RunTemplate runTempalte = (RunTemplate) ele;
                String tagName = runTempalte.getTagName();
                if (tagName.matches("^\\w+$") && !resultMap.containsKey(tagName)) {
                    resultMap.put(tagName, defaultPlaceholder);
                }
            } else if (ele instanceof IterableTemplate) {
                IterableTemplate iterableTemplate = (IterableTemplate) ele;
                initFieldByDocument(iterableTemplate.getTemplates(), resultMap, defaultPlaceholder);
            } else if (ele instanceof BlockTemplate) {
                BlockTemplate blockTemplate = (BlockTemplate) ele;
                initFieldByDocument(blockTemplate.getTemplates(), resultMap, defaultPlaceholder);
            }
        });
    }

    /**
     * 删除表格行测试（最简单版本）
     * 目前：只要是这一行有跨列的则不删除这个单元格
     * TODO 如果跨行，又怎么处理？？
     * @throws Exception
     */
    @Test
    public void removeLine() throws Exception {
        String resource = "src/test/resources/template/grid_bu.docx";
        Map<String, Object> data = new HashMap<>();
        data.put("r34", 50);
        data.put("empty", null);
        data.put("rs1_show", true);
        data.put("rs2_show", null);
        data.put("rs4_show", null);
        data.put("rs5_show", null);
        data.put("rs6_show", null);
        data.put("rs7_show", null);
        Configure config = Configure.builder()
            .addPlugin('$', new RemoveTableRowRenderPolicy("——"))
            .useSpringEL(false)
            .build();
        XWPFTemplate template = XWPFTemplate.compile(resource, config);
        initFieldByDocument(template.getElementTemplates(), data, "——");
        template.render(data);
        template.writeToFile("target/out_grid_bu.docx");

        resource = "src/test/resources/template/delete_row.docx";
        XWPFTemplate template2 = XWPFTemplate.compile(resource, config);
        initFieldByDocument(template2.getElementTemplates(), data, "——");
        template2.render(data);
        template2.writeToFile("target/out_remove_line.docx");
    }
}
