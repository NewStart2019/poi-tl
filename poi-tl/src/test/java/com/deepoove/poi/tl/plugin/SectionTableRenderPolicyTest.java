package com.deepoove.poi.tl.plugin;

import java.util.*;

import com.deepoove.poi.plugin.table.*;
import com.deepoove.poi.template.BlockTemplate;
import com.deepoove.poi.template.IterableTemplate;
import com.deepoove.poi.template.MetaTemplate;
import com.deepoove.poi.template.run.RunTemplate;
import com.deepoove.poi.util.WordTableUtils;
import org.junit.jupiter.api.Test;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;

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
     * @throws Exception
     */
    @Test
    public void removeLine() throws Exception {
        String resource = "src/test/resources/template/grid_bu.docx";
        Map<String, Object> data = new HashMap<>();
        data.put("r34", 50);
        data.put("empty", null);
        data.put("rs1_show", null);
        data.put("rs2_show", null);
        data.put("rs4_show", true);
        data.put("rs5_show", true);
        data.put("rs6_show", null);
        data.put("rs7_show", false);
        data.put("open_addRow", true);
        data.put("rs2_show_insertPosition", 4);
        data.put("test", "C");
        Configure config = Configure.builder()
            .addPlugin('$', new RemoveTableRowRenderPolicy("——"))
            .useSpringEL(false)
            .build();
        resource = "src/test/resources/template/delete_row.docx";
        XWPFTemplate template2 = XWPFTemplate.compile(resource, config);
        initFieldByDocument(template2.getElementTemplates(), data, "——");
        template2.render(data);
        template2.writeToFile("target/out_remove_line.docx");
    }

    public Map<String, Object> init3(int number) {
        Map<String, Object> test = new HashMap<>();
        test.put("companyName", "测试公司");
        test.put("org_email", "4398430@ee.com");
        test.put("org_queryPhone", "56486");
        test.put("org_address", "56486");
        test.put("is_check", "56486");
        test.put("conclusion", "符合");
        List<Map<String, Object>> data = new ArrayList<>();
        test.put("subRecords", data);
        test.put("subRecords_number", 29);
        test.put("subRecords_reduce", 0);
        Random random = new Random();
        for (int i = 1; i <= number; i++) {
            Map<String, Object> e1 = new HashMap<>();
            data.add(e1);
            e1.put("sjbh1", random.nextInt(1000));
            e1.put("sjbh2", random.nextInt(1000));
            e1.put("sjbh3", random.nextInt(1000));
            e1.put("lq", i);
            e1.put("jcbw1", "检测部位" + i);
            e1.put("rq", "技术指标" + i);
            e1.put("item", "混凝土抗折" + i);
            e1.put("L1", 30);
            e1.put("L2", 10);
            e1.put("L3", 20);
            e1.put("p1", 20);
        }
        return test;
    }

    /**
     * 本示例是 测试的多行表格渲染策略 和 删除表格策略同时 存在的场景测试
     */
    @Test
    public void tesRemoveTable() throws Exception {
        // 测试支持多行表头和单行表头
        ArrayList<Integer> conditions = new ArrayList<>();
        resource = "src/test/resources/template/mutiple_row_table.docx";
        conditions.add(3);
        conditions.add(5);
        conditions.add(8);
        conditions.add(14);
        conditions.add(20);
        conditions.add(23);
        conditions.add(80);
        LoopRowTableAllRenderPolicy policy = new LoopRowTableAllRenderPolicy();
        for (Integer condition : conditions) {
            Map<String, Object> stringObjectMap = init3(condition);
            // Map<String, Object> stringObjectMap = init2(50);
            stringObjectMap.put("subRecords_rendermode", 7);
            stringObjectMap.put("subRecords_row_number", 3);
            stringObjectMap.put("subRecords_first_number", 15);
            stringObjectMap.put("subRecords_number", 27);
            stringObjectMap.put("subRecords_mode", 2);
            stringObjectMap.put("blank_desc", "以下空白");
            stringObjectMap.put("blank_vmerge", "以下空白");
            stringObjectMap.put("deleteTable", condition % 3 == 0);
            Configure config = Configure.builder()
                .useSpringEL(false)
                .addPlugin('-', new RemoveTableRenderPolicy())
                .bind("subRecords", policy)
                .bind("deleteTable", new RemoveTableRenderPolicy())
                .build();
            XWPFTemplate template = XWPFTemplate.compile(resource, config).render(stringObjectMap);
            WordTableUtils.setMinHeightParagraph(template.getXWPFDocument());
            template.writeToFile("target/out_mutiple_row_table" + condition + ".docx");
        }
    }

}
