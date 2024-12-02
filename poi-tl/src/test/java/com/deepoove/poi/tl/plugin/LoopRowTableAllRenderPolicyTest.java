package com.deepoove.poi.tl.plugin;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import com.deepoove.poi.data.Pictures;
import com.deepoove.poi.plugin.table.LoopFullTableInsertFillRenderPolicy;
import com.deepoove.poi.plugin.table.LoopRowTableAllRenderPolicy;
import com.deepoove.poi.plugin.table.LoopRowTableRenderPolicy;
import com.deepoove.poi.util.WordTableUtils;
import com.deepoove.poi.xwpf.NiceXWPFDocument;
import org.apache.poi.xwpf.usermodel.*;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectType;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STSectionMark;

import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.*;

@DisplayName("Example for HackLoop Table")
public class LoopRowTableAllRenderPolicyTest {

    String resource = "src/test/resources/template/render_hackloop.docx";
    PaymentHackData data = new PaymentHackData();
    LoopRowTableAllRenderPolicy policy = new LoopRowTableAllRenderPolicy();

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
    public void testDefaultLoopTablePolicyExample() throws Exception {
        LoopRowTableRenderPolicy hackLoopSameLineTableRenderPolicy = new LoopRowTableRenderPolicy(true);
        Configure config = Configure.builder()
            .useSpringEL(false)
            .bind("goods", policy)
            .bind("labors", policy)
            .bind("goods2", hackLoopSameLineTableRenderPolicy)
            .bind("labors2", hackLoopSameLineTableRenderPolicy)
            .build();
        XWPFTemplate template = XWPFTemplate.compile(resource, config).render(data);
        WordTableUtils.setMinHeightParagraph(template.getXWPFDocument());
        template.writeToFile("target/out_table_render_row_span.docx");
    }

    public Map<String, Object> init2(int number) {
        Map<String, Object> test = new HashMap<>();
        test.put("companyName", "测试公司");
        test.put("org_email", "4398430@ee.com");
        test.put("org_queryPhone", "56486");
        test.put("org_address", "56486");
        List<Map<String, Object>> data = new ArrayList<>();
        test.put("test", data);
        test.put("test_number", 29);
        test.put("test_reduce", 0);
        test.put("conclusion", "结论");
        for (int i = 1; i <= number; i++) {
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
        LoopRowTableAllRenderPolicy loopRowTableAllRenderPolicy = new LoopRowTableAllRenderPolicy(false, true);
        resource = "src/test/resources/template/render_existed_fill.docx";
        ArrayList<Integer> conditions = new ArrayList<>();
        conditions.add(10);
        conditions.add(65);
        for (Integer condition : conditions) {
            Map<String, Object> stringObjectMap = init2(condition);
            stringObjectMap.put("test_rendermode", 1);
            Configure config = Configure.builder()
                .useSpringEL(false)
                .bind("test", policy)
                .build();
            XWPFTemplate template = XWPFTemplate.compile(resource, config).render(stringObjectMap);
            WordTableUtils.removeLastBlankParagraph(template.getXWPFDocument());
            WordTableUtils.setMinHeightParagraph(template.getXWPFDocument());
            template.writeToFile("target/out_existed" + condition + ".docx");
        }
    }

    @Test
    public void testLoopExistedAndFillBlanRow() throws Exception {
        resource = "src/test/resources/template/render_existed_fill.docx";
        ArrayList<Integer> conditions = new ArrayList<>();
        conditions.add(10);
        conditions.add(27);
        conditions.add(56);
        conditions.add(65);
        conditions.add(200);
        for (Integer condition : conditions) {
            Map<String, Object> stringObjectMap = init2(condition);
            stringObjectMap.put("test_rendermode", 2);
            stringObjectMap.put("test_mode", 3);
            // stringObjectMap.put("test_nofill", 2);
            policy.setSaveNextLine(true);
            Configure config = Configure.builder()
                .useSpringEL(false)
                .bind("test", policy)
                .build();
            XWPFTemplate template = XWPFTemplate.compile(resource, config).render(stringObjectMap);
            WordTableUtils.removeLastBlankParagraph(template.getXWPFDocument());
            WordTableUtils.setMinHeightParagraph(template.getXWPFDocument());
            template.writeToFile("target/out_exiest_fill" + condition + ".docx");
        }
    }

    @Test
    public void testLoopFillRow() throws Exception {
        policy = new LoopRowTableAllRenderPolicy(false, true);
        resource = "src/test/resources/template/render_insert_fill_3.docx";
        ArrayList<Integer> conditions = new ArrayList<>();
        conditions.add(10);
        conditions.add(21);
        conditions.add(23);
        conditions.add(25);
        conditions.add(40);
        conditions.add(45);
        conditions.add(48);
        conditions.add(49);
        conditions.add(50);
        conditions.add(69);
        conditions.add(73);
        conditions.add(75);
        conditions.add(77);
        conditions.add(80);
        for (Integer condition : conditions) {
            Map<String, Object> stringObjectMap = init2(condition);
            stringObjectMap.put("test_number", 21);
            // stringObjectMap.put("test_reduce", 1);
            // stringObjectMap.put("test_nofill", 1);
            stringObjectMap.put("test_mode", 2);
            stringObjectMap.put("test_header", 3);
            stringObjectMap.put("test_footer", 4);
            stringObjectMap.put("blank_desc", "以下空白");
            stringObjectMap.put("test_rendermode", 3);
            Configure config = Configure.builder()
                .useSpringEL(false)
                .bind("test", policy)
                .build();
            XWPFTemplate template = XWPFTemplate.compile(resource, config).render(stringObjectMap);
            WordTableUtils.removeLastBlankParagraph(template.getXWPFDocument());
            WordTableUtils.setMinHeightParagraph(template.getXWPFDocument());
            template.writeToFile("target/out_insert_fill" + condition + ".docx");
        }
    }

    @Test
    public void testLoopFullTableRow() throws Exception {
        LoopFullTableInsertFillRenderPolicy hackLoopTableRenderPolicy2 = new LoopFullTableInsertFillRenderPolicy(false);
        resource = "src/test/resources/template/render_insert_fill_mutiple_template.docx";
        ArrayList<Integer> conditions = new ArrayList<>();
        conditions.add(0);
        conditions.add(11);
        conditions.add(12);
        conditions.add(15);
        conditions.add(23);
        conditions.add(24);
        conditions.add(25);
        conditions.add(30);
        for (Integer condition : conditions) {
            Map<String, Object> stringObjectMap = init2(condition);
            stringObjectMap.put("test_number", 24);
            stringObjectMap.put("test_mode", 2);
            stringObjectMap.put("test_rendermode", 4);
            stringObjectMap.put("test_row_number", 2);
            stringObjectMap.put("test_vmerge", 2);
            // stringObjectMap.put("test_reduce", 1);
            // stringObjectMap.put("test_remove_next_line", 4);
            stringObjectMap.put("blank_desc", "以下空白");
            Configure config = Configure.builder()
                .useSpringEL(false)
                .bind("test", policy)
                .build();
            XWPFTemplate template = XWPFTemplate.compile(resource, config).render(stringObjectMap);
            WordTableUtils.setMinHeightParagraph(template.getXWPFDocument());
            NiceXWPFDocument xwpfDocument = template.getXWPFDocument();
            // 最后几行内容为空，删除元素避免产生新页
            WordTableUtils.removeLastBlankParagraph(xwpfDocument);
            WordTableUtils.setMinHeightParagraph(xwpfDocument);
            template.writeToFile("target/out_loop_full_table" + condition + ".docx");
        }
    }

    public Map<String, Object> init3(int first, int second) {
        Map<String, Object> test = new HashMap<>();
        test.put("companyName", "测试公司");
        test.put("org_email", "439828430@ee.com");
        test.put("org_queryPhone", "56486");
        test.put("org_address", "56486");
        test.put("conclusion", "符合");
        List<Map<String, Object>> data = new ArrayList<>();
        test.put("test", data);
        test.put("test_number", 29);
        test.put("test_reduce", 0);

        for (int f = 0; f < first; f++) {
            Map<String, Object> fMap = new HashMap<>();
            data.add(fMap);
            fMap.put("conclusion", "你自己弄吧" + f);
            List<Map<String, Object>> subs = new ArrayList<>();
            fMap.put("subs", subs);
            for (int i = 1; i <= second; i++) {
                Map<String, Object> e1 = new HashMap<>();
                subs.add(e1);
                e1.put("xh", i);
                e1.put("qywz", "测试位置" + i);
                e1.put("rq", "技术指标" + i);
                e1.put("jcjg", "检测结果" + i);
                e1.put("jgpd", "结果判定" + i);
                e1.put("a", 10);
                e1.put("b", 20);
            }
        }
        return test;
    }


    @Test
    public void testLoopFullTableIncludeSubTableRow() throws Exception {
        // resource = "src/test/resources/template/render_insert_fill.docx";
        resource = "src/test/resources/template/render_insert_fill_mutiple_template.docx";
        ArrayList<Integer> conditions = new ArrayList<>();
        conditions.add(0);
        conditions.add(10);
        conditions.add(25);
        conditions.add(30);
        conditions.add(50);
        conditions.add(60);
        conditions.add(80);
        for (Integer condition : conditions) {
            Map<String, Object> stringObjectMap;
            if (condition == 0) {
                stringObjectMap = init3(0, condition);
            } else {
                stringObjectMap = init3(3, condition);
            }
            stringObjectMap.put("test_rendermode", 5);
            stringObjectMap.put("test_number", 24);
            stringObjectMap.put("test_mode", 3);
            stringObjectMap.put("test_row_number",
                resource.contains("render_insert_fill_mutiple_template") ? 2 : 1);
            // stringObjectMap.put("test_vmerge", 2);
            // stringObjectMap.put("test_remove_next_line", 4);
            stringObjectMap.put("blank_desc", "以下空白");
            Configure config = Configure.builder()
                .useSpringEL(false)
                .bind("test", policy)
                .build();
            XWPFTemplate template = XWPFTemplate.compile(resource, config).render(stringObjectMap);
            WordTableUtils.removeLastBlankParagraph(template.getXWPFDocument());
            WordTableUtils.setMinHeightParagraph(template.getXWPFDocument());
            template.writeToFile("target/out_loop_sub_table" + condition + ".docx");
        }
    }

    @Test
    public void testLoopCopyHeaderRowRenderPolicy() throws Exception {
        // 测试支持多行表头和单行表头
        resource = "src/test/resources/template/render_insert_fill_2.docx";
        ArrayList<Integer> conditions = new ArrayList<>();
        conditions.add(10);
        conditions.add(24);
        conditions.add(30);
        conditions.add(52);
        conditions.add(60);
        conditions.add(80);
        for (Integer condition : conditions) {
            Map<String, Object> stringObjectMap = init2(condition);
            stringObjectMap.put("test_first_number", 24);
            stringObjectMap.put("test_number", 28);
            stringObjectMap.put("test_mode", 1);
            stringObjectMap.put("test_rendermode", 6);
            stringObjectMap.put("test_remove_next_line", 4);
            stringObjectMap.put("blank_desc", "以下空白");
            Configure config = Configure.builder()
                .useSpringEL(false)
                .bind("test", policy)
                .build();
            XWPFTemplate template = XWPFTemplate.compile(resource, config).render(stringObjectMap);
            WordTableUtils.setMinHeightParagraph(template.getXWPFDocument());
            template.writeToFile("target/out_loop_copy_header" + condition + ".docx");
        }
    }

    public Map<String, Object> init3(int number) {
        Map<String, Object> test = new HashMap<>();
        test.put("companyName", "测试公司");
        test.put("org_email", "4398430@ee.com");
        test.put("org_queryPhone", "56486");
        test.put("org_address", "56486");
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

    @Test
    public void testLoopMutipleRow() throws Exception {
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
        for (Integer condition : conditions) {
            Map<String, Object> stringObjectMap = init3(condition);
            // Map<String, Object> stringObjectMap = init2(50);
            stringObjectMap.put("subRecords_rendermode", 7);
            stringObjectMap.put("subRecords_row_number", 3);
            stringObjectMap.put("subRecords_first_number", 15);
            stringObjectMap.put("subRecords_number", 27);
            stringObjectMap.put("subRecords_mode", 2);
            stringObjectMap.put("blank_desc", "以下空白");
            Configure config = Configure.builder()
                .useSpringEL(false)
                .bind("subRecords", policy)
                .build();
            XWPFTemplate template = XWPFTemplate.compile(resource, config).render(stringObjectMap);
            WordTableUtils.setMinHeightParagraph(template.getXWPFDocument());
            template.writeToFile("target/out_mutiple_row_table" + condition + ".docx");
        }
    }

    public Map<String, Object> init4(int number) {
        Map<String, Object> test = new HashMap<>();
        test.put("report_no", "5864986054");
        test.put("projectCode", "439843054");
        List<Map<String, Object>> data = new ArrayList<>();
        test.put("subRecords", data);
        test.put("subRecords2", data);
        test.put("subRecords_number", 29);
        test.put("subRecords_reduce", 0);
        Random random = new Random();
        for (int i = 1; i <= number; i++) {
            Map<String, Object> e1 = new HashMap<>();
            data.add(e1);
            e1.put("jc1", i);
            e1.put("wg1", "检测部位" + i);
            e1.put("wg2", "技术指标" + i);
            e1.put("wg3", "混凝土抗折" + i);
            e1.put("jg1", 30);
            e1.put("jg2", 10);
            e1.put("jg3", 20);
            e1.put("zt1", 20);
            e1.put("fc1", "随便做");
            e1.put("fc2", "随便做");
            e1.put("fc3", "随便做");
            e1.put("bw1", "部位1_" + i);
            e1.put("bw2", "部位2_" + i);
            e1.put("bw3", "部位3_" + i);
            e1.put("image_base64", Pictures.of("src/test/resources/picture/p.png").create());
            e1.put("image2_base64", Pictures.of("src/test/resources/picture/p.png").create());
            e1.put("image3_base64", Pictures.of("src/test/resources/picture/p.png").create());
        }
        return test;
    }

    @Test
    public void testLoopMutipleRowIncludePicture() throws Exception {
        // 测试支持多行表头和单行表头
        ArrayList<Integer> conditions = new ArrayList<>();
        resource = "src/test/resources/util/double_mutiple_copy_picture.docx";
        // resource = "D:\\DingTalkAppData\\DingTalk\\download\\外墙节能构造及保温层厚度（钻芯法）检测报告.docx";
        conditions.add(8);
        conditions.add(13);
        conditions.add(20);
        conditions.add(23);
        conditions.add(80);
        for (Integer condition : conditions) {
            Map<String, Object> stringObjectMap = init4(condition);
            stringObjectMap.put("subRecords_rendermode", 7);
            stringObjectMap.put("subRecords_row_number", 3);
            stringObjectMap.put("subRecords_first_number", 6);
            stringObjectMap.put("subRecords_number", 6);
            stringObjectMap.put("subRecords_mode", 2);
            stringObjectMap.put("subRecords_fpdb", 2);

            stringObjectMap.put("subRecords2_rendermode", 7);
            stringObjectMap.put("subRecords2_row_number", 3);
            stringObjectMap.put("subRecords2_first_number", 3);
            stringObjectMap.put("subRecords2_number", 3);
            stringObjectMap.put("subRecords2_mode", 2);
            stringObjectMap.put("subRecords2_fpdb", 2);
            stringObjectMap.put("blank_desc", "以下空白");
            Configure config = Configure.builder()
                .useSpringEL(false)
                .bind("subRecords", policy)
                .bind("subRecords2", policy)
                .build();
            XWPFTemplate template = XWPFTemplate.compile(resource, config).render(stringObjectMap);
            WordTableUtils.setMinHeightParagraph(template.getXWPFDocument());
            template.writeToFile("target/out_double_mutiple_copy_picture" + condition + ".docx");
        }
    }

    public Map<String, Object> init5(int number) {
        Map<String, Object> test = new HashMap<>();
        test.put("companyName", "测试公司");
        test.put("org_email", "4398430@ee.com");
        test.put("fillPersonName", "李四");
        test.put("create_time", "56486");
        test.put("receivePerson", "李世明");
        test.put("receiveTime", "2024-11-21 10:00:00");
        test.put("signedTime", "2024-11-21 12:00:00");
        test.put("qrcode", Pictures.of("src/test/resources/picture/p.png").create());
        test.put("signedPersonQrcode", Pictures.of("src/test/resources/picture/p.png").create());
        List<Map<String, Object>> data = new ArrayList<>();
        test.put("subRecords", data);
        test.put("subRecords_number", 29);
        test.put("subRecords_reduce", 0);
        Random random = new Random();
        for (int i = 1; i <= number; i++) {
            Map<String, Object> e1 = new HashMap<>();
            data.add(e1);
            e1.put("wt_no", 100 + random.nextInt(10));
            if (i == 3) {
                continue;
            }
            e1.put("inspeItemName", "检测项目" + i);
            e1.put("wtrq", "2024-11-21 10:00:00");
            e1.put("report_no", "19849884" + i);
            e1.put("jcEgName", "测试工程");
            e1.put("remark", "无");
            e1.put("p1", 20);
        }
        return test;
    }


    @Test
    public void testLoopMutilpleRowRenderSaveSuffixPolicy() throws Exception {
        // 测试支持多行表头和单行表头
        ArrayList<Integer> conditions = new ArrayList<>();
        resource = "src/test/resources/util/mutiple_suffix.docx";
        conditions.add(5);
        conditions.add(9);
        conditions.add(10);
        conditions.add(13);
        conditions.add(20);
        conditions.add(22);
        conditions.add(24);
        conditions.add(26);
        conditions.add(30);
        conditions.add(80);
        for (Integer condition : conditions) {
            Map<String, Object> stringObjectMap = init5(condition);
            stringObjectMap.put("subRecords_rendermode", 8);
            stringObjectMap.put("subRecords_row_number", 1);
            stringObjectMap.put("subRecords_first_number", 13);
            stringObjectMap.put("subRecords_number", 13);
            stringObjectMap.put("subRecords_mode", 3);
            stringObjectMap.put("subRecords_external_footer", 4);
            // stringObjectMap.put("subRecords_nofill", 0);
            stringObjectMap.put("blank_desc", "以下空白");
            Configure config = Configure.builder()
                .useSpringEL(false)
                .bind("subRecords", policy)
                .build();
            XWPFTemplate template = XWPFTemplate.compile(resource, config).render(stringObjectMap);
            WordTableUtils.setMinHeightParagraph(template.getXWPFDocument());
            template.writeToFile("target/out_mutiple_suffix" + condition + ".docx");
        }
    }

    @Test
    void testDeleteTable() throws IOException {
        resource = "D:\\DingTalkAppData\\DingTalk\\download\\保温层厚度修订版.docx";
        NiceXWPFDocument niceXWPFDocument = new NiceXWPFDocument(Files.newInputStream(Paths.get(resource)));

        // 遍历所有段落
        for (IBodyElement element : niceXWPFDocument.getBodyElements()) {
            if (element.getElementType() == BodyElementType.PARAGRAPH) {
                XWPFParagraph paragraph = (XWPFParagraph) element;
                System.out.println(paragraph.getText());
                // 如果段落包含分节符，则表示这一节结束
                if (paragraph.getCTP().getPPr() != null && paragraph.getCTP().getPPr().getSectPr() != null) {
                    CTSectPr sectPr = paragraph.getCTP().getPPr().getSectPr();
                    CTSectType type = sectPr.getType();
                    System.out.println("------ Section Break ------");
                }
            } else if (element.getElementType() == BodyElementType.TABLE) {
                XWPFTable table = (XWPFTable) element;
                System.out.println(table.getText());
            } else if (element.getElementType() == BodyElementType.CONTENTCONTROL) {
                XWPFSDT sdt = (XWPFSDT) element;
                System.out.println(sdt.getContent().getText());
            }
        }

        // 创建一个新的Word文档
        XWPFDocument document = new XWPFDocument();

        // 添加一些文本
        XWPFParagraph paragraph = document.createParagraph();
        XWPFRun run = paragraph.createRun();
        run.setText("这是第一部分的内容。");

        // 在这里插入分节符
        WordTableUtils.setSectionBreak(document, STSectionMark.NEXT_PAGE, null);

        // 继续添加更多文本
        paragraph = document.createParagraph();
        run = paragraph.createRun();
        run.setText("这是第二部分的内容。");

        // 保存文档
        try (FileOutputStream out = new FileOutputStream("C:\\Users\\Administrator\\Desktop\\b.docx")) {
            document.write(out);
        }

        // 关闭文档
        document.close();
        // niceXWPFDocument.write(new FileOutputStream("C:\\Users\\Administrator\\Desktop\\a.docx"));
    }
}
