package com.deepoove.poi.tl.util;

import com.deepoove.poi.util.WordTableUtils;
import org.apache.poi.wp.usermodel.HeaderFooterType;
import org.apache.poi.xwpf.usermodel.*;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

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
        XWPFTable newTable = WordTableUtils.copyTable(document, table);

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
    void testCopyTableRow() throws Exception {
        String file = "src/test/resources/template/render_insert_fill.docx";

        try (FileInputStream fileInputStream = new FileInputStream(file);
             XWPFDocument document = new XWPFDocument(fileInputStream)) {
            XWPFTable table = document.getTables().get(0);
            // 创建的表格默认有1行1列
            XWPFTable table2 = document.createTable();
            WordTableUtils.copyLineContent(table.getRow(0), table2.insertNewTableRow(0), 0);
            WordTableUtils.copyLineContent(table.getRow(1), table2.insertNewTableRow(1), 1);
            WordTableUtils.removeLastRow(table2);
            WordTableUtils.copyTableTblPr(table, table2);
            table2.getCTTbl().getTblGrid();
            // 断言跨列属性是否复制过来
            assertTrue(table2.getRow(1).getCell(2).getCTTc().getTcPr().isSetVMerge());
            assertTrue(table2.getRow(1).getCell(3).getCTTc().getTcPr().isSetVMerge());
            out_file = "target/copy_line_test.docx";
            // 保存文档
            try (FileOutputStream out = new FileOutputStream(out_file)) {
                document.write(out);
            }
        }
    }
}
