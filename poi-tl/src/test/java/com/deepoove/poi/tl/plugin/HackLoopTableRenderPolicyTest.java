package com.deepoove.poi.tl.plugin;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import com.deepoove.poi.data.Pictures;
import com.deepoove.poi.plugin.table.LoopExistedAndFillRowTableRenderPolicy;
import com.deepoove.poi.plugin.table.LoopExistedRowTableRenderPolicy;
import com.deepoove.poi.plugin.table.LoopRowTableAndFillRenderPolicy;
import com.deepoove.poi.plugin.table.LoopRowTableRenderPolicy;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.util.*;

@DisplayName("Example for HackLoop Table")
public class HackLoopTableRenderPolicyTest {

    String resource = "src/test/resources/template/render_hackloop.docx";
    PaymentHackData data = new PaymentHackData();

    @BeforeEach
    public void init() {
        List<Goods> goods = new ArrayList<>();
        Goods good = new Goods();
        good.setCount(4);
        good.setName("墙纸");
        good.setDesc("书房卧室");
        good.setDiscount(1500);
        good.setPrice(400);
        good.setTax(new Random().nextInt(10) + 20);
        good.setTotalPrice(1600);
        good.setPicture(Pictures.ofLocal("src/test/resources/earth.png").size(24, 24).create());
        good.setTotal("1024");
        for (int i = 0; i < 4; i++) {
            goods.add(good);
        }
        data.setGoods(goods);

        List<Labor> labors = new ArrayList<>();
        Labor labor = new Labor();
        labor.setCategory("油漆工");
        labor.setPeople(2);
        labor.setPrice(400);
        labor.setTotalPrice(1600);
        labors.add(labor);
        labors.add(labor);
        labors.add(labor);
        data.setLabors(labors);

        data.setTotal("1024");

        // same line
        data.setGoods2(goods);
        data.setLabors2(labors);

    }

    @Test
    public void testPaymentHackExample() throws Exception {
        LoopRowTableRenderPolicy hackLoopTableRenderPolicy = new LoopRowTableRenderPolicy();
        LoopRowTableRenderPolicy hackLoopSameLineTableRenderPolicy = new LoopRowTableRenderPolicy(true);
        Configure config = Configure.builder().bind("goods", hackLoopTableRenderPolicy)
            .bind("labors", hackLoopTableRenderPolicy).bind("goods2", hackLoopSameLineTableRenderPolicy)
            .bind("labors2", hackLoopSameLineTableRenderPolicy).build();
        XWPFTemplate template = XWPFTemplate.compile(resource, config).render(data);
        template.writeToFile("target/out_table_render_row_span.docx");
    }

    public Map<String, Object> init2() {
        Map<String, Object> test = new HashMap<>();
        test.put("companyName", "测试公司");
        List<Map<String, Object>> data = new ArrayList<>();
        test.put("test", data);
        test.put("testnumber", 29);
        test.put("testreduce", 0);
        for (int i = 1; i <= 65; i++) {
            Map<String, Object> e1 = new HashMap<>();
            data.add(e1);
            e1.put("xh", i);
            e1.put("qywz", "测试位置" + i);
            e1.put("rq", "技术指标" + i);
            e1.put("jcjg", "检测结果" + i);
            e1.put("jgpd", "结果判定" + i);
            e1.put("a", 10);
            e1.put("b", 20);
        }
        return test;
    }

    @Test
    public void testLoopExistedRow() throws Exception {
        LoopExistedRowTableRenderPolicy hackLoopTableRenderPolicy = new LoopExistedRowTableRenderPolicy(false, true);
        resource = "D:\\DingTalkAppData\\DingTalk\\download\\路基路面几何尺寸（宽度）试验检测报告.docx";
        resource = "src/test/resources/template/render_existed_fill.docx";
        Map<String, Object> stringObjectMap = init2();
        Configure config = Configure.builder()
            .bind("test", hackLoopTableRenderPolicy)
            .build();
        XWPFTemplate template = XWPFTemplate.compile(resource, config).render(stringObjectMap);
        template.writeToFile("target/out_existed.docx");
    }

    @Test
    public void testLoopExistedAndFillBlanRow() throws Exception {
        LoopExistedAndFillRowTableRenderPolicy hackLoopTableRenderPolicy2 = new LoopExistedAndFillRowTableRenderPolicy(false, true);
        resource = "src/test/resources/template/render_existed_fill.docx";
        Map<String, Object> stringObjectMap = init2();
        Configure config = Configure.builder()
            .bind("test", hackLoopTableRenderPolicy2)
            .build();
        XWPFTemplate template = XWPFTemplate.compile(resource, config).render(stringObjectMap);
        template.writeToFile("target/out_exiest_fill.docx");
    }

    @Test
    public void testLoopFillRow() throws Exception {
        LoopRowTableAndFillRenderPolicy hackLoopTableRenderPolicy2 = new LoopRowTableAndFillRenderPolicy(false, true);
        resource = "src/test/resources/template/render_insert_fill.docx";
        Map<String, Object> stringObjectMap = init2();
        stringObjectMap.put("testnumber", 29);
        stringObjectMap.put("testreduce", 0);
        stringObjectMap.put("testmode", 2);
        stringObjectMap.put("testheader", 1);
        stringObjectMap.put("testfooter", 4);
        stringObjectMap.put("blank_desc", "以下空白");
        Configure config = Configure.builder()
            .bind("test", hackLoopTableRenderPolicy2)
            .build();
        XWPFTemplate template = XWPFTemplate.compile(resource, config).render(stringObjectMap);
        template.writeToFile("target/out_insert_fill.docx");
    }



}
