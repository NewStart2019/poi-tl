package com.deepoove.poi.tl.plugin;

import java.util.ArrayList;
import java.util.List;
import java.util.Random;

import com.deepoove.poi.plugin.table.LoopExistedRowTableRenderPolicy;
import com.deepoove.poi.plugin.table.RemoveTableRowRenderPolicy;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import com.deepoove.poi.data.Pictures;
import com.deepoove.poi.plugin.table.LoopRowTableRenderPolicy;

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
        goods.add(good);
        goods.add(good);
        goods.add(good);
        goods.add(good);
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
        template.writeToFile("target/out_render_looprow.docx");
    }

    @Test
    public void testLoopExistedRow() throws Exception {
        LoopExistedRowTableRenderPolicy hackLoopTableRenderPolicy = new LoopExistedRowTableRenderPolicy(false, true);
        Configure config = Configure.builder()
            .bind("goods", hackLoopTableRenderPolicy)
//            .bind("labors", hackLoopTableRenderPolicy)
            .build();
        XWPFTemplate template = XWPFTemplate.compile(resource, config).render(data);
        template.writeToFile("target/out_render_loopexistedrow.docx");
    }

}
