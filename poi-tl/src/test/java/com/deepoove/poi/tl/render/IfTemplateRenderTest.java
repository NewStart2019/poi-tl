package com.deepoove.poi.tl.render;

import static org.junit.jupiter.api.Assertions.assertEquals;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import com.deepoove.poi.config.Configure;
import com.deepoove.poi.xwpf.NiceXWPFDocument;
import org.apache.poi.xwpf.usermodel.IBody;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.tl.source.XWPFTestSupport;

@DisplayName("If template test case")
public class IfTemplateRenderTest {

    @SuppressWarnings("serial")
    @Test
    public void testIfFalse() throws Exception {
        Map<String, Object> datas = new HashMap<String, Object>() {
            {
                put("title", "poi-tl");
                put("isShowTitle", true);
                put("showUser", false);
                put("showDate", false);
            }
        };

        XWPFTemplate template = XWPFTemplate.compile("src/test/resources/template/iterable_if1.docx");
        template.render(datas);

        XWPFDocument document = XWPFTestSupport.readNewDocument(template);
        XWPFParagraph paragraph = document.getParagraphArray(0);
        assertEquals(paragraph.getText(), "Hi, poi-tl");

        XWPFTable table = document.getTableArray(0);
        XWPFTableCell cell = table.getRow(1).getCell(0);
        assertEquals(cell.getText(), "Hi, poi-tl");

        XWPFHeader header = document.getHeaderArray(0);
        paragraph = header.getParagraphArray(0);
        assertEquals(paragraph.getText(), "Hi, poi-tl");
    }

    @SuppressWarnings("serial")
    @Test
    public void testIfTrue() throws Exception {
        Map<String, Object> datas = new HashMap<String, Object>() {
            {
                put("title", "poi-tl");
                put("isShowTitle", true);
                put("showUser", new HashMap<String, Object>() {
                    {
                        put("user", "Sayi");
                        put("showDate", new HashMap<String, Object>() {
                            {
                                put("date", "2020-02-10");
                            }
                        });
                    }
                });
            }
        };

        XWPFTemplate template = XWPFTemplate.compile("src/test/resources/template/iterable_if1.docx");
        template.render(datas);

        XWPFDocument document = XWPFTestSupport.readNewDocument(template);
        XWPFTable table = document.getTableArray(0);
        XWPFTableCell cell = table.getRow(1).getCell(0);
        XWPFHeader header = document.getHeaderArray(0);

        testParagraph(document);
        testParagraph(cell);
        testParagraph(header);
    }

    private void testParagraph(IBody document) {
        XWPFParagraph paragraph = document.getParagraphArray(0);
        assertEquals(paragraph.getText(), "Hi, poi-tl");

        paragraph = document.getParagraphArray(1);
        assertEquals(paragraph.getText(), "Hello, My perfect.");

        paragraph = document.getParagraphArray(2);
        assertEquals(paragraph.getText(), "UserName: Sayi");

        paragraph = document.getParagraphArray(3);
        assertEquals(paragraph.getText(), "Date: 2020-02-10");

        paragraph = document.getParagraphArray(4);
        assertEquals(paragraph.getText(), "Date: 2020-02-10");

        paragraph = document.getParagraphArray(5);
        assertEquals(paragraph.getText(), "I love this Game Date: 2020-02-10 and good game.");
    }

    @SuppressWarnings("serial")
    @Test
    public void testBasicIf() throws Exception {
        Map<String, Object> datas = new HashMap<String, Object>() {
            {
                put("title", "poi-tl");
                put("show", true);
            }
        };

        XWPFTemplate template = XWPFTemplate.compile("src/test/resources/template/iterable_if2.docx");
        template.render(datas);
        template.writeToFile("target/out_iterable_if_basic.docx");
    }

    @SuppressWarnings("serial")
    @Test
    public void testBasicIf3() throws Exception {
        List<Map<String, Object>> list = new ArrayList<>();
        for (int i = 16; i <= 22; i++) {
            Map<String, Object> a = new HashMap<>();
            list.add(a);
            a.put("a", i);
            a.put("a" + i, i);
            a.put("name", "李四" + i);
            a.put("age", i + "岁");
        }
        Map<String, Object> datas = new HashMap<String, Object>() {
            {
                put("title", "poi-tl");
                put("show", true);
                put("data", list);
                put("name", "纵横四海");
                put("age", "100");
            }
        };
        Configure config = Configure.builder().useSpringEL(false).build();
        XWPFTemplate template = XWPFTemplate.compile("src/test/resources/template/iterable_if3.docx", config);
        template.render(datas);
        NiceXWPFDocument xwpfDocument = template.getXWPFDocument();
        template.writeToFile("target/out_iterable_if_basic.docx");
    }

}
