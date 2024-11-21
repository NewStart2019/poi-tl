package com.deepoove.poi.tl.util;

import com.deepoove.poi.util.UnitUtils;
import com.deepoove.poi.util.WordTableUtils;
import com.deepoove.poi.xwpf.NiceXWPFDocument;
import org.apache.logging.log4j.util.Strings;
import org.apache.poi.wp.usermodel.HeaderFooterType;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlObject;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.springframework.util.Assert;

import javax.xml.namespace.QName;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;

import static org.junit.jupiter.api.Assertions.assertEquals;

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
        WordTableUtils.mergeMutipleLine(table, 0, 2);

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
        FileInputStream fileInputStream = new FileInputStream(file);
        XWPFDocument document = new XWPFDocument(fileInputStream);
        XWPFTable table = document.getTables().get(1);

        WordTableUtils.mergeMutipleLine(table, 3, 4);
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
        String file = "src/test/resources/template/render_insert_fill.docx";
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

        XWPFDocument document2 = new XWPFDocument(new FileInputStream(out_file));
        assertEquals(2, document2.getTables().size());
    }

    @Test
    void testfindVMerge() throws Exception {
        String file = "src/test/resources/template/render_insert_fill_2.docx";
        try (FileInputStream fileInputStream = new FileInputStream(file);
             XWPFDocument document = new XWPFDocument(fileInputStream)) {
            XWPFTable table = document.getTables().get(0);
            int verticalMergedRows = WordTableUtils.findVerticalMergedRows(table, 0, 0);
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

            XWPFParagraph paragraph = document.createParagraph();
            XWPFRun run = paragraph.createRun();

            run.setText("这是第一行文本");
            run.addBreak(BreakType.COLUMN); // 添加换行
            run.setText("这是同一段落中的第二行文本");

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
            XWPFTable table1 = document.getTables().get(0);
            XWPFTable table2 = document.getTables().get(1);
            XWPFParagraph paragraphArray1 = document.getParagraphArray(3);

            int posOfTable = document.getPosOfTable(table2);
            WordTableUtils.setElementPostion(document, table2, document.getPosOfTable(table1));
            WordTableUtils.setElementPostion(document, table1, posOfTable);
            WordTableUtils.setElementPostion(document, paragraphArray1, document.getBodyElements().size() - 1);
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
    void testSetVMerge() throws IOException {
        String template = "src/test/resources/template/render_insert_fill_2.docx";
        FileInputStream fileInputStream = new FileInputStream(template);
        NiceXWPFDocument document = new NiceXWPFDocument(fileInputStream);
        XWPFTable xwpfTable = document.getTables().get(1);
        CTTc ctTc = xwpfTable.getRow(0).getCell(1).getCTTc();
        System.out.println(ctTc.getTcPr().getVMerge().getVal());
        Assert.isTrue(ctTc.getTcPr().getVMerge().getVal() == STMerge.RESTART, "跨列设置成功");
        document.write(Files.newOutputStream(Paths.get("target/out_copy_border.docx")));
    }
}
