package com.deepoove.poi.tl.util;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import com.deepoove.poi.data.style.Style;
import com.deepoove.poi.plugin.table.LoopRowTableAllRenderPolicy;
import com.deepoove.poi.util.StyleUtils;
import com.deepoove.poi.util.UnitUtils;
import com.deepoove.poi.util.word.WordTableUtils;
import com.deepoove.poi.xwpf.NiceXWPFDocument;
import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.logging.log4j.util.Strings;
import org.apache.poi.wp.usermodel.HeaderFooterType;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlObject;
import org.junit.jupiter.api.Disabled;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import javax.xml.namespace.QName;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;
import static org.junit.jupiter.api.Assertions.assertFalse;

class WordTableUtilsTest {

    String out_file = "target/table_utils.docx";

    @Test
    void testParagraphCopy() throws Exception {
        String template = "src/test/resources/util/copy_template.docx";
        String template2 = "target/out_copy_template_copys.docx";
        try (FileInputStream fileInputStream = new FileInputStream(template);
             XWPFDocument document = new XWPFDocument(fileInputStream);
             XWPFDocument document2 = new XWPFDocument()) {
            XWPFTable table1 = document.getTables().get(0);
            XWPFParagraph oldParagraph = document.getParagraphs().get(1);

            // 测试带样式和不带样式复制段落
            WordTableUtils.copyParagraph(oldParagraph,
                document.createParagraph(), false);
            WordTableUtils.copyParagraph(oldParagraph,
                document.createParagraph(), true);
            // 从XWPFDocument段落复制到 单元格段落中
            WordTableUtils.copyParagraph(oldParagraph,
                table1.getRow(1).getCell(1).addParagraph(), true);

            // 复制到页眉中去
            XWPFHeader header = document2.createHeader(HeaderFooterType.FIRST);
            XWPFParagraph paragraph = header.createParagraph();
            WordTableUtils.copyParagraph(oldParagraph, paragraph, true);

            // 测试是否支持跨表格复制段落
            WordTableUtils.copyParagraph(oldParagraph,
                document2.createParagraph(), true);
            XWPFTable table = document2.createTable(1, 1);
            WordTableUtils.setTableWidthA4(table);
            WordTableUtils.copyParagraph(oldParagraph, table.getRow(0).getCell(0).addParagraph(), true);

            XWPFTable table2 = document.getTables().get(0);

            XWPFParagraph pictureParagraph = document.getParagraphs().get(2);
            WordTableUtils.copyParagraph(pictureParagraph, document2.createParagraph(), true);

            out_file = "target/out_copy_template.docx";
            // 保存文档
            try (FileOutputStream out = new FileOutputStream(out_file);
                 FileOutputStream out2 = new FileOutputStream(template2)) {
                document.write(out);
                document2.write(out2);
            }
        }
    }

    @Test
    void testCellCopy() throws Exception {
        String template = "src/test/resources/util/copy_template.docx";
        String template2 = "target/out_copy_cell_copys.docx";
        try (FileInputStream fileInputStream = new FileInputStream(template);
             XWPFDocument document = new XWPFDocument(fileInputStream);
             XWPFDocument document2 = new XWPFDocument()) {
            XWPFTable table1 = document.getTables().get(0);
            XWPFTableCell table1Cell = table1.getRow(0).getCell(0);
            XWPFTable table2 = document.getTables().get(1);
            XWPFTableCell table2Cell = table2.getRow(0).getCell(0);
            XWPFTableCell table2Cell2 = table2.getRow(1).getCell(2);
            WordTableUtils.copyCell(table1Cell, table2Cell, false);
            WordTableUtils.copyCell(table1Cell, table2Cell2, true);

            // 跨文档
            XWPFTable table = document2.createTable(2, 1);
            WordTableUtils.setTableWidthA4(table);
            WordTableUtils.copyCell(table1Cell, table.getRow(0).getCell(0), false);
            WordTableUtils.copyCell(table1Cell, table.getRow(1).getCell(0), true);

            out_file = "target/out_copy_cell.docx";
            // 保存文档
            try (FileOutputStream out = new FileOutputStream(out_file);
                 FileOutputStream out2 = new FileOutputStream(template2)) {
                document.write(out);
                document2.write(out2);
            }
        }
    }

    @Test
    void testRowCopy() throws Exception {
        String template = "src/test/resources/util/copy_template.docx";
        String template2 = "target/out_copy_row_copys.docx";
        try (FileInputStream fileInputStream = new FileInputStream(template);
             XWPFDocument document = new XWPFDocument(fileInputStream);
             XWPFDocument document2 = new XWPFDocument()) {
            XWPFTable table1 = document.getTables().get(0);


            // 跨文档
            XWPFTable table = document2.createTable();
            WordTableUtils.setTableWidthA4(table);
            WordTableUtils.removeLastRow(table);

            out_file = "target/out_copy_row.docx";
            // 保存文档
            try (FileOutputStream out = new FileOutputStream(out_file);
                 FileOutputStream out2 = new FileOutputStream(template2)) {
                document.write(out);
                document2.write(out2);
            }
        }
    }

    @Test
    void mergeRowAndWriteSlash() throws IOException {
        // 创建一个新的 Word 文档
        XWPFDocument document = new XWPFDocument();
        out_file = "target/out_merege_line.docx";

        XWPFTable table = getXwpfTable(document);

        // 合并多行
        WordTableUtils.mergeMultipleLine(table, 0, 2);

        // 设置对角线
        XWPFTableCell cellRow00 = table.getRow(0).getCell(0);
        WordTableUtils.setDiagonalBorder(cellRow00);

        try (FileOutputStream out = new FileOutputStream(out_file)) {
            document.write(out);
        }
        document.close();
        System.out.println("文档已保存，包含斜切效果的形状。");
    }

    private static XWPFTable getXwpfTable(XWPFDocument document) {
        // 创建一个 2 行 3 列的表格
        XWPFTable table = document.createTable(3, 3);
        WordTableUtils.setTableWidthA4(table);

        // 向第一行的单元格写入一些文本
        XWPFTableRow row1 = table.getRow(0);
        row1.getCell(0).setText("单元格 1");
        row1.getCell(1).setText("单元格 2");
        row1.getCell(2).setText("单元格 3");

        // 向第二行的单元格写入一些文本
        XWPFTableRow row2 = table.getRow(1);
        row2.getCell(0).setText("单元格 4");
        row2.getCell(1).setText("单元格 5");
        row2.getCell(2).setText("单元格 6");

        XWPFTableRow row3 = table.getRow(2);
        row3.getCell(0).setText("单元格 4");
        row3.getCell(1).setText("单元格 5");
        row3.getCell(2).setText("单元格 6");
        return table;
    }

    @Test
    void testMergeMutipleLineIncludeVMerge() throws Exception {
        // 创建一个新的 Word 文档
        String file = "src/test/resources/template/iterable_payment.docx";
        file = "C:\\Users\\Administrator\\Desktop\\报告邮寄(1).docx";
        FileInputStream fileInputStream = new FileInputStream(file);
        XWPFDocument document = new XWPFDocument(fileInputStream);
        XWPFTable table = document.getTables().get(0);

        WordTableUtils.mergeMultipleLine(table, 2, 5);
        out_file = "target/out_merged_table.docx";
        // 保存文档
        try (FileOutputStream out = new FileOutputStream(out_file)) {
            document.write(out);
        }
        document.close();
    }

    @Test
    void testCopyTable() throws Exception {
        // 创建一个新的 Word 文档
        String file = "src/test/resources/util/copy_template.docx";
        FileInputStream fileInputStream = new FileInputStream(file);
        XWPFDocument document = new XWPFDocument(fileInputStream);
        XWPFTable table = document.getTables().get(0);
        XWPFTable newTable = WordTableUtils.copyTable(document, table, true);

        out_file = "target/test.docx";
        // 保存文档
        try (FileOutputStream out = new FileOutputStream(out_file)) {
            document.write(out);
        }
        document.close();
    }

    @Test
    void testfindVMerge() throws Exception {
        String file = "src/test/resources/template/render_insert_fill_2.docx";
        try (FileInputStream fileInputStream = new FileInputStream(file);
             XWPFDocument document = new XWPFDocument(fileInputStream)) {
            XWPFTable table = document.getTables().get(0);
            int verticalMergedRows = WordTableUtils.findVerticalMergedRows(table, 0, 0);
            assertEquals(3, verticalMergedRows);

            verticalMergedRows = WordTableUtils.findVerticalMergedRows(table, 0, 2);
            assertEquals(3, verticalMergedRows);

            verticalMergedRows = WordTableUtils.findVerticalMergedRows(table, 0, 3);
            assertEquals(3, verticalMergedRows);

            verticalMergedRows = WordTableUtils.findVerticalMergedRows(table, 0, 4);
            assertEquals(1, verticalMergedRows);

            verticalMergedRows = WordTableUtils.findVerticalMergedRows(table, 0, 5);
            assertEquals(1, verticalMergedRows);

            verticalMergedRows = WordTableUtils.findVerticalMergedRows(table, 0, 6);
            assertEquals(3, verticalMergedRows);
        }
    }

    @Test
    void testEnter() throws Exception {
        out_file = "target/test.docx";
        try (XWPFDocument document = new XWPFDocument()) {
            // 写入换行：创建一个段落，自动另起一行
            // 创建第一个段落并添加文本
            // XWPFParagraph paragraph1 = document.createParagraph();   //******
            // XWPFRun run1 = paragraph1.createRun();
            // run1.setText("这是第一行文本");

            // 创建第二个段落，这将自动产生一个换行效果
            // XWPFParagraph paragraph2 = document.createParagraph();   //******
            // XWPFRun run2 = paragraph2.createRun();
            // run2.setText("这是第二行文本");

            // 2. 非结构化文本 - 性别复选框（使用 □ ☑）
            XWPFParagraph genderParagraph = document.createParagraph();
            genderParagraph.createRun().setText("Select Gender:");

            // 男性复选框
            XWPFRun maleCheckbox = genderParagraph.createRun();
            maleCheckbox.setText("\u25A1"); // 未选中
            // maleCheckbox.setFontSize(12); // 设置字体大小
            maleCheckbox.setFontFamily("Wingdings 2"); // 确保字体支持符号

            // 女性复选框
            XWPFRun femaleCheckbox = genderParagraph.createRun();
            // femaleCheckbox.setFontSize(12); // 设置字体大小
            // femaleCheckbox.setText("☑ Female");
            femaleCheckbox.setText("R"); // 选中
            femaleCheckbox.setFontFamily("Wingdings 2"); // 确保字体支持符号

            // 简单结构化文本
            {
                XWPFParagraph paragraph = document.createParagraph();
                XWPFRun run = paragraph.createRun();
                run.setText("Below is a structured document tag:");

                // 创建 SDT (结构化文档标签块)
                CTSdtRun sdtRun = paragraph.getCTP().addNewSdt();
                CTSdtPr sdtPr = sdtRun.addNewSdtPr();
                CTString alias = sdtPr.addNewAlias();
                alias.setVal("Sample SDT");

                CTSdtContentRun sdtContent = sdtRun.addNewSdtContent();
                CTR ctr = sdtContent.addNewR();
                CTText ctText = ctr.addNewT();
                ctText.setStringValue("This is SDT content.");
            }


            // 1. 富文本（Rich Text）
            {
                XWPFParagraph richTextParagraph = document.createParagraph();
                XWPFRun richTextRun = richTextParagraph.createRun();
                richTextRun.setText("This is Complex Rich Text SDT:");

                CTSdtRun sdtRichText = richTextParagraph.getCTP().addNewSdt();
                CTSdtPr sdtPr2 = sdtRichText.addNewSdtPr();
                sdtPr2.addNewAlias().setVal("Complex Rich Text");

                // 设置富文本内容
                CTSdtContentRun sdtContent2 = sdtRichText.addNewSdtContent();

                XWPFRun contentRun1 = richTextParagraph.createRun();
                contentRun1.setText("Bold Text");
                contentRun1.setBold(true);

                XWPFRun contentRun2 = richTextParagraph.createRun();
                contentRun2.setText(" | Italic Text");
                contentRun2.setItalic(true);

                XWPFRun contentRun3 = richTextParagraph.createRun();
                contentRun3.setText(" | Underlined Text");
                contentRun3.setUnderline(UnderlinePatterns.SINGLE);

                XWPFRun contentRun4 = richTextParagraph.createRun();
                contentRun4.setText(" | Colored Text");
                contentRun4.setColor("FF5733");

                // 将富文本内容添加到SDT中
                sdtContent2.getDomNode().appendChild(contentRun1.getCTR().getDomNode().cloneNode(true));
                sdtContent2.getDomNode().appendChild(contentRun2.getCTR().getDomNode().cloneNode(true));
                sdtContent2.getDomNode().appendChild(contentRun3.getCTR().getDomNode().cloneNode(true));
                sdtContent2.getDomNode().appendChild(contentRun4.getCTR().getDomNode().cloneNode(true));
            }

            // 2. 日期选择器
            {
                XWPFParagraph dateParagraph = document.createParagraph();
                XWPFRun dateRun = dateParagraph.createRun();
                dateRun.setText("This is Date Picker SDT:");

                CTSdtRun sdtDate = dateParagraph.getCTP().addNewSdt();
                sdtDate.addNewSdtContent().addNewR().addNewT().setStringValue("2025-05-08");
            }


            // 3. 下拉列表
            {
                // 在文档中创建一个新的段落。
                XWPFParagraph dropdownParagraph = document.createParagraph();
                // 在段落中创建一个新的 SDT
                CTSdtRun sdtDropdown = dropdownParagraph.getCTP().addNewSdt();
                // 添加 SDT 属性（CTSdtPr）
                CTSdtPr dropdownPr = sdtDropdown.addNewSdtPr();
                // 为 SDT 设置一个别名（Alias）
                dropdownPr.addNewAlias().setVal("Multi-Option Dropdown");

                // 创建数据绑定（绑定到 XML 数据源）: 这个数据绑定通常用于高级需求，绑定到外部 XML 数据。
                CTDataBinding dataBinding = dropdownPr.addNewDataBinding();
                // 定义 XPath 路径，指向下拉选项数据。
                dataBinding.setXpath("/options/option");

                // 设置下拉框选项: 这里添加了三个选项：Option 1, Option 2, Option 3。
                CTSdtDropDownList ctSdtDropDownList = dropdownPr.addNewDropDownList();
                CTSdtListItem listItem1 = ctSdtDropDownList.addNewListItem();
                listItem1.setDisplayText("Option 1");
                listItem1.setValue("Option 1");

                CTSdtListItem listItem2 = ctSdtDropDownList.addNewListItem();
                listItem2.setDisplayText("Option 2");
                listItem2.setValue("Option 2");

                CTSdtListItem listItem3 = ctSdtDropDownList.addNewListItem();
                listItem3.setDisplayText("Option 3");
                listItem3.setValue("Option 3");

                // 设置下拉框默认值
                CTSdtContentRun ctSdtContentRun = sdtDropdown.addNewSdtContent();
                CTR ctr = ctSdtContentRun.addNewR();
                CTText ctText = ctr.addNewT();
                ctText.setStringValue("Option 1");
            }

            // 4. 复选框
            {
                XWPFParagraph checkboxParagraph = document.createParagraph();
                CTSdtRun sdtCheckbox1 = checkboxParagraph.getCTP().addNewSdt();

                // 设置复选框文本
                CTSdtContentRun contentRun3 = sdtCheckbox1.addNewSdtContent();
                CTR ctr3 = contentRun3.addNewR();
                CTText ctText3 = ctr3.addNewT();
                ctText3.setStringValue("☑ 男 \u25A1 女");
            }

            // 5. 单选框
            {
                XWPFParagraph radioParagraph = document.createParagraph();
                XWPFRun radioRun = radioParagraph.createRun();
                radioRun.setText("This is Radio Button SDT:");

                CTSdtRun sdtRadio = radioParagraph.getCTP().addNewSdt();
                sdtRadio.addNewSdtContent().addNewR().addNewT().setStringValue("◉ Radio");
            }

            try (FileOutputStream out = new FileOutputStream(out_file)) {
                document.write(out);
            }
        }
    }

    @Test
    void testReduceHeight() throws Exception {
        out_file = "target/test.docx";
        try (XWPFDocument document = new XWPFDocument()) {
            XWPFTable table = document.createTable(1, 1);
            XWPFTableCell cell = table.getRow(0).getCell(0);
            XWPFParagraph paragraph = cell.getParagraphs().get(0);
            XWPFRun run = paragraph.createRun();
            run.setText("这是第一行文本");
            WordTableUtils.setTableRowHeight(table.getRow(0), UnitUtils.point2Twips(24), STHeightRule.EXACT);
            Double fontSizeAsDouble = run.getFontSizeAsDouble();
            for (int i = 1; i < 29; i++) {
                WordTableUtils.copyLineContent(table.getRow(0), table.insertNewTableRow(i), i);
            }
            int rowHeight = WordTableUtils.findRowHeight(table.getRow(0));
            ruduceRowHeigth(table, 0, -1);
            try (FileOutputStream out = new FileOutputStream(out_file)) {
                document.write(out);
            }
        }
    }

    public static void ruduceRowHeigth(XWPFTable table, int startIndex, int endIndex) {
        if (endIndex == -1) {
            endIndex = table.getRows().size() - 1;
        }
        int rowNumber = endIndex - startIndex + 1;
        int tableMargin = WordTableUtils.findTableMargin(table, 2);
        // 默认行距：如果不手动设置，XWPFParagraph的行距是单倍行距，具体数值取决于Word应用的默认设置
        // 240：表示1倍行距
        int sum = tableMargin + UnitUtils.point2Twips(24);
        int perRowReduce = sum / rowNumber;
        int remain = sum % rowNumber;
        // perRowReduce += (remain == 0 ? 0 : 1);
        for (int i = startIndex; i <= endIndex; i++) {
            XWPFTableRow row = table.getRow(i);
            int rowHeight = WordTableUtils.findRowHeight(row);
            WordTableUtils.setTableRowHeight(row, rowHeight - perRowReduce, STHeightRule.EXACT);
        }
        for (int i = endIndex - remain + 1; i <= endIndex; i++) {
            XWPFTableRow row = table.getRow(i);
            int rowHeight = WordTableUtils.findRowHeight(row);
            WordTableUtils.setTableRowHeight(row, rowHeight - 1, STHeightRule.AT_LEAST);
        }
    }

    @Test
    void testChangeElement() {
        String template = "src/test/resources/util/copy_template.docx";
        try (FileInputStream fileInputStream = new FileInputStream(template);
             XWPFDocument document = new XWPFDocument(fileInputStream)) {
            XWPFTable table2 = document.getTables().get(1);
            WordTableUtils.setElementPosition(document, table2, 0);
            IBodyElement iBodyElement = document.getBodyElements().get(4);
            WordTableUtils.setElementPosition(document, iBodyElement, 0);
            try (FileOutputStream out = new FileOutputStream(out_file)) {
                document.write(out);
            }
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    @Test
    void testXmlCursorMove() {
        String template = "src/test/resources/util/copy_template.docx";
        try (FileInputStream fileInputStream = new FileInputStream(template);
             XWPFDocument document = new XWPFDocument(fileInputStream)) {
            // XmlCursor cursor = document.getDocument().getBody().newCursor();
            String xmlString = "<root>\n" +
                "    <parent>\n" +
                "        <child>文本1</child>\n" +
                "        <child>文本2</child>\n" +
                "    </parent>\n" +
                "    <sibling>文本3</sibling>\n" +
                "</root>";
            XmlObject xmlObject = XmlObject.Factory.parse(xmlString);
            XmlCursor cursor = xmlObject.newCursor();
            cursor.toStartDoc();  // 将光标移动到文档的起始位置

            while (cursor.hasNextToken()) {
                cursor.toNextToken();
                XmlCursor.TokenType tokenType = cursor.currentTokenType();
                System.out.println(String.format("当前节点类型-名称：%s-%s", tokenType, cursor.getName()));
                if (tokenType.isText() && Strings.isNotBlank(cursor.getChars())) {
                    System.out.println("当前文本内容：" + cursor.getChars());
                }
            }

            cursor.toStartDoc();
            // 定位到 <root> 元素
            cursor.selectPath("$this//root");
            cursor.toFirstChild();
            cursor.toFirstChild();

            // Start a new <work> element
            cursor.beginElement("work");
            cursor.insertChars("555-555-5555");
            // cursor.toNextToken();

            cursor.toStartDoc();
            // 定位到 <parent> 元素
            cursor.selectPath("$this/root/parent");
            cursor.toNextSelection();
            cursor.toFirstChild(); // 定位到 <parent> 的开始位置

            // 使用 beginElement() 插入 <newChild> 标签对
            cursor.beginElement("newChild");  // 插入 <newChild> 开始标签，并将光标移动到其内部
            // 注意这里插入 属性的顺序不能 放在 插入内容之后，否则报错: Can only insert attributes before other attributes or after containers.
            cursor.insertAttributeWithValue("attr", "attrValue");
            cursor.insertChars("New Content");  // 插入内容到 <newChild> 内部
            cursor.toParent();  // 返回到 <parent> 标签

            QName qname = new QName("http://2006", "test", "v");
            // 插入一个新元素，且只能插入这样插入值，否则在标签外面，需要进入标签里面插入值
            cursor.insertElementWithText(qname, "New Element Text");

            System.out.println("Modified XML:\n" + xmlObject.xmlText());
            cursor.close();  // 释放资源
        } catch (IOException | XmlException e) {
            throw new RuntimeException(e);
        }
    }

    @Test
    void testXmlCursorCopy() {
        try {
            // XmlCursor cursor = document.getDocument().getBody().newCursor();
            String xmlString = "<root>\n" +
                "    <parent>\n" +
                "        <child>文本1</child>\n" +
                "        <child>文本2</child>\n" +
                "    </parent>\n" +
                "    <sibling>文本3</sibling>\n" +
                "</root>";

            XmlObject xmlObject = XmlObject.Factory.parse(xmlString);
            // 创建光标定位到要移动的标签位置（第一个 <child>）
            XmlCursor sourceCursor = xmlObject.newCursor();
            sourceCursor.selectPath("$this//child");
            sourceCursor.toNextSelection();  // 定位到第一个 <child> 标签

            // 创建目标光标，定位到 <sibling> 标签的开始位置
            XmlCursor targetCursor = xmlObject.newCursor();
            targetCursor.selectPath("$this//sibling");
            targetCursor.toNextSelection();
            // targetCursor.toStartToken();  // 确保移动到 <sibling> 的开始位置

            // 移动 <child> 标签到 <sibling> 标签之前
            sourceCursor.moveXml(targetCursor);

            // 打印结果
            System.out.println("Modified XML:\n" + xmlObject.xmlText());

            // 清理资源
            sourceCursor.close();
            targetCursor.close();
        } catch (XmlException e) {
            throw new RuntimeException(e);
        }
    }

    @Test
    void testCopyLeftBorder() throws IOException {
        String template = "src/test/resources/util//copy_border.docx";
        FileInputStream fileInputStream = new FileInputStream(template);
        NiceXWPFDocument document = new NiceXWPFDocument(fileInputStream);
        XWPFTable xwpfTable = document.getTables().get(0);
        WordTableUtils.setBottomBorder(xwpfTable, null);
        document.write(Files.newOutputStream(Paths.get("target/out_copy_border.docx")));
    }

    @Test
    void testCalculateFontSize() throws IOException {
        // 使用 src/test/resources 下的模板，不再依赖本机桌面绝对路径。
        String template = "src/test/resources/util/copy_border.docx";
        try (FileInputStream in = new FileInputStream(template);
             NiceXWPFDocument document = new NiceXWPFDocument(in)) {
            XWPFTable table = document.getTables().get(0);
            XWPFTableRow row = table.getRow(0);
            XWPFTableCell cell = row.getCell(0);
            int cellWidth = cell.getWidth();
            int cellHeight = WordTableUtils.findCellHeight(cell);
            // 基准字号优先取模板 run 的原始字号；未显式设置时才回退默认字号 12。
            XWPFRun existingRun = cell.getParagraphs().get(0).getRuns().get(0);
            double baseSize = expectedBaseSize(existingRun);

            // 1) 长文本在有限高度内放不下：字号应从基准字号向下缩小，且仍在可读范围内。
            String longText = "铜遂人才共育园（科能技校园）项目（一期用地一期建设）施工铜遂人才共育园（科能技校园）项目";
            double shrunk = WordTableUtils.adjustFontSize(cell, longText, cellWidth, cellHeight);
            assertTrue(shrunk <= baseSize, "长文本字号应小于基准字号");
            assertTrue(shrunk >= 1, "字号不应小于最小可读字号");

            // 2) 短文本放得下：保持基准字号。
            assertEquals(baseSize, WordTableUtils.adjustFontSize(cell, "短", cellWidth, cellHeight),
                "短文本应保持基准字号");

            // 3) null / 空内容不抛异常，返回基准字号。
            assertEquals(baseSize, WordTableUtils.adjustFontSize(cell, null, cellWidth, cellHeight));
            assertEquals(baseSize, WordTableUtils.adjustFontSize(cell, "", cellWidth, cellHeight));

            // 4) 非正宽高不做缩放，返回基准字号（避免除零退化）。
            assertEquals(baseSize, WordTableUtils.adjustFontSize(cell, longText, 0, cellHeight));
            assertEquals(baseSize, WordTableUtils.adjustFontSize(cell, longText, cellWidth, 0));

            // 5) 应用计算出的字号并写出，确认可正常生成文档。
            XWPFParagraph paragraph = cell.getParagraphs().get(0);
            XWPFRun run = paragraph.getRuns().isEmpty() ? paragraph.createRun()
                : paragraph.getRuns().get(0);
            Style style = Style.builder().buildFontSize(shrunk).build();
            run.setText(longText, 0);
            StyleUtils.styleRun(run, style);
            try (FileOutputStream out = new FileOutputStream("target/adjusted-table.docx")) {
                document.write(out);
            }
        }
    }

    @Test
    void testCalculateFontSize2() throws IOException {
        // re_FSP0_DLZM01_subRecords
        LoopRowTableAllRenderPolicy policy = new LoopRowTableAllRenderPolicy();
        String resource = "src/test/resources/template/道路照明（现用）.docx";
        // 先读取 JSON 文件内容为字符串（UTF-8），再用 Jackson 转换为 Map。
        // 嵌套对象映射为 Map、数组映射为 List，数字/字符串/布尔分别映射为对应 Java 类型。
        String json = Files.readString(
            Paths.get("src/test/resources/data/adjustFontSize.json"));
        Map<String, Object> stringObjectMap = new ObjectMapper().readValue(
            json, new TypeReference<Map<String, Object>>() {});
        Configure config = Configure.builder()
            .useSpringEL(false)
            .addPlugin('*', new FitCellPolicy())
            // {{re_FSQ0_SWRW01_subRecords}}序号
            .bind("re_FSP0_DLZM01_subRecords", policy)
            .build();
        XWPFTemplate template = XWPFTemplate.compile(resource, config).render(stringObjectMap);
        WordTableUtils.setMinHeightParagraph(template.getXWPFDocument());
        WordTableUtils.removeLastBlankParagraph(template.getXWPFDocument());
        WordTableUtils.setMinHeightParagraph(template.getXWPFDocument());
        template.writeToFile("target/out_道路照明.docx");
    }

    /**
     * 计算预期基准字号：优先取 run 显式字号（即使大于默认 12），缺失或非法时回退默认字号 12。
     * 与 {@link WordTableUtils#adjustFontSize} 的基准字号口径保持一致。
     */
    private static double expectedBaseSize(XWPFRun run) {
        double size = WordTableUtils.findFontSize(run);
        return size <= 0 ? 12 : size;
    }

    /**
     * 构造内存单元格；fontSize 小于等于 0 表示不设置显式字号（用于验证回退默认字号）。
     */
    private static XWPFTableCell newMemoryCell(XWPFDocument document, double fontSize) {
        XWPFTableCell cell = document.createTable(1, 1).getRow(0).getCell(0);
        XWPFRun run = cell.getParagraphs().get(0).createRun();
        if (fontSize > 0) {
            run.setFontSize((int) fontSize);
        }
        run.setText("x");
        return cell;
    }

    @Test
    void testBoldItalicStyleCombination() {
        // 粗体与斜体同时设置时应能共存（位或组合），方法不应抛异常并返回有效字号。
        XWPFDocument document = new XWPFDocument();
        XWPFTable table = document.createTable(1, 1);
        XWPFTableCell cell = table.getRow(0).getCell(0);
        XWPFRun run = cell.getParagraphs().get(0).createRun();
        run.setFontSize(14);
        run.setBold(true);
        run.setItalic(true);
        assertTrue(run.isBold());
        assertTrue(run.isItalic());
        double size = WordTableUtils.adjustFontSize(cell, "粗斜体文本内容", 1200, 600);
        assertTrue(size >= 1.0 && size <= 14.0);

        // 显式关闭（w:val=0）：isBold/isItalic 应返回 false。
        XWPFRun run2 = cell.getParagraphs().get(0).createRun();
        run2.setBold(false);
        run2.setItalic(false);
        assertFalse(run2.isBold());
        assertFalse(run2.isItalic());
    }

    @Test
    void testFindCellWidthAndHeightWithMergedCells() throws Exception {
        // 1) 垂直合并（vMerge）：restart 与 continuation 都应返回合并后的总高。
        try (XWPFDocument doc = new XWPFDocument(
            new FileInputStream("src/test/resources/util/self_define_vmerge.docx"))) {
            XWPFTable table = doc.getTables().get(0);
            // 普通单元格：高度即所在行高度。
            assertEquals(table.getRow(0).getHeight(),
                WordTableUtils.findCellHeight(table.getRow(0).getCell(0)));
            // r1 起始格（restart）：行高 618 + 下一行 617 = 1235。
            assertEquals(1235, WordTableUtils.findCellHeight(table.getRow(1).getCell(0)));
            // r2 后续格（continuation）：同样返回合并总高 1235。
            assertEquals(1235, WordTableUtils.findCellHeight(table.getRow(2).getCell(0)));
            // 水平合并：r0c3 gridSpan=5，宽度为 5 个网格列之和 3633。
            assertEquals(3633, WordTableUtils.findCellWidth(table.getRow(0).getCell(3)));
        }

        // 2) 同时含 gridSpan 与裸 vMerge（无 val，按 OOXML 语义默认 continue）的复杂模板。
        try (XWPFDocument doc = new XWPFDocument(
            new FileInputStream("src/test/resources/template/mutiple_row_table.docx"))) {
            XWPFTable table = doc.getTables().get(0);
            // r4c1 跨两列，宽度=网格列 262+561=823。
            assertEquals(823, WordTableUtils.findCellWidth(table.getRow(4).getCell(1)));
            // r4c1 同时为垂直合并起始，跨 3 行，高度=454×3=1362。
            assertEquals(1362, WordTableUtils.findCellHeight(table.getRow(4).getCell(1)));
            // 裸 vMerge 后续格（旧 findVMerge 会误判为 0）仍应返回合并总高。
            assertEquals(1362, WordTableUtils.findCellHeight(table.getRow(5).getCell(1)));
        }

        // 3) 空值防护：不抛异常，返回 0。
        assertEquals(0, WordTableUtils.findCellWidth(null));
        assertEquals(0, WordTableUtils.findCellHeight(null));
    }
}
