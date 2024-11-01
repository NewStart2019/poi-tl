package com.deepoove.poi.tl.xwpf;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.Map;

import com.deepoove.poi.util.WordTableUtils;
import org.apache.poi.ooxml.POIXMLProperties.CoreProperties;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFSDT;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlObject;
import org.junit.jupiter.api.Test;

import com.deepoove.poi.XWPFTemplate;

/**
 * @author Sayi
 */
public class SDTTest {

//    paragraph.getIRuns().stream().filter(r -> r instanceof XWPFSDT).forEach(r -> {
//        ISDTContent isdtContent = ((XWPFSDT) r).getContent();
//        if (isdtContent instanceof XWPFSDTContent) {
//            @SuppressWarnings("unchecked")
//            List<ISDTContents> contents = (List<ISDTContents>) ReflectionUtils.getValue("bodyElements",
//                    isdtContent);
//            List<XWPFRun> collect = contents.stream()
//                    .filter(c -> c instanceof XWPFRun)
//                    .map(c -> (XWPFRun) c)
//                    .collect(Collectors.toList());
//            // to do refactor sdtcontent
//            resolveXWPFRuns(collect, metaTemplates, stack);
//        }
//    });

    @SuppressWarnings("serial")
    @Test
    public void testRenderSDTInParagraph() throws Exception {
        Map<String, Object> data = new HashMap<String, Object>() {
            {
                put("titlefd", "Poi-tl");
                put("name", "模板引擎");
                put("list", new ArrayList<Map<String, Object>>() {
                    {
                        add(Collections.singletonMap("name", "Lucy"));
                        add(Collections.singletonMap("name", "Hanmeimei"));
                    }
                });
            }
        };

        XWPFTemplate.compile("src/test/resources/template/template_sdt.docx")
            .render(data)
            .writeToFile("target/out_sdt_para.docx");

    }

    @Test
    public void testRenderSDTBlockInBody() throws Exception {
        @SuppressWarnings("serial")
        Map<String, Object> data = new HashMap<String, Object>() {
            {
                put("title", "Poi-tl");
                put("name", "模板引擎");
                put("list", new ArrayList<Map<String, Object>>() {
                    {
                        add(Collections.singletonMap("name", "Lucy"));
                        add(Collections.singletonMap("name", "Hanmeimei"));
                    }
                });
            }
        };

        XWPFTemplate.compile("src/test/resources/template/sdt.docx")
            .render(data)
            .writeToFile("target/out_sdt_block.docx");

    }

    @Test
    public void testRenderSDTInTextbox() throws Exception {
        @SuppressWarnings("serial")
        Map<String, Object> data = new HashMap<String, Object>() {
            {
                put("A", "Poi-tl");
            }
        };

        XWPFTemplate template = XWPFTemplate.compile("src/test/resources/template/sdt_core.docx").render(data);
        CoreProperties coreProperties = template.getXWPFDocument().getProperties().getCoreProperties();
        coreProperties.setSubjectProperty("Poi-tl手册");
        template.writeToFile("target/out_sdt_core.docx");

    }

    @Test
    public void testRenderSDTInTableRow() throws Exception {
        @SuppressWarnings("serial")
        Map<String, Object> data = new HashMap<String, Object>() {
            {
                put("title", "Poi-tl");
            }
        };

        XWPFTemplate template = XWPFTemplate.compile("src/test/resources/template/sdt_cell.docx").render(data);
        CoreProperties coreProperties = template.getXWPFDocument().getProperties().getCoreProperties();
        coreProperties.setTitle("poi-tl");
        coreProperties.setDescription("desc");
        ;
        template.writeToFile("target/out_sdt_cell.docx");

    }

    @Test
    void testBreak() throws IOException {
        try (FileInputStream fis = new FileInputStream("src/test/resources/template/insert_paragraph.docx");
             XWPFDocument document = new XWPFDocument(fis)) {
            WordTableUtils.setPageBreak(document,  document.getTables().get(0));

            // 保存文档
            try (FileOutputStream fos = new FileOutputStream("target/out_insert_paragraph.docx")) {
                document.write(fos);
            }
            System.out.println("New paragraph inserted successfully!");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}
