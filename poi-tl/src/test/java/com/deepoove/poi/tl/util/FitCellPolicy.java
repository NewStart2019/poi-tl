package com.deepoove.poi.tl.util;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.policy.RenderPolicy;
import com.deepoove.poi.template.ElementTemplate;
import com.deepoove.poi.template.run.RunTemplate;
import com.deepoove.poi.util.word.WordTableUtils;
import org.apache.poi.xwpf.usermodel.LineSpacingRule;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;

public class FitCellPolicy implements RenderPolicy {
    @Override
    public void render(ElementTemplate eleTemplate, Object data, XWPFTemplate template) {
        XWPFRun run = ((RunTemplate) eleTemplate).getRun();
        String thing = data == null ? "" : String.valueOf(data);
        try {
            XWPFTableCell cell = ((XWPFTableCell) ((XWPFParagraph) run.getParent()).getBody());
            int width = cell.getWidth();
            int height = WordTableUtils.findCellHeight(cell);
            XWPFParagraph xwpfParagraph = ((XWPFParagraph) run.getParent());
            xwpfParagraph.setSpacingBetween(1, LineSpacingRule.AUTO);
            double defaultFontSize = WordTableUtils.findFontSize(run);
            if (defaultFontSize == 0) {
                defaultFontSize = 10.5;
            }
            // 计算适合的字体大小
            double fontSize = WordTableUtils.adjustFontSize(cell, thing, width, height);
            run.setFontSize(Math.max(fontSize, 6));
            run.setText(thing, 0);
        } catch (Exception ex) {
            try {
                run.setText(thing, 0);
            } catch (Exception ex1) {
                ex1.printStackTrace();
            }
        }
    }
}
