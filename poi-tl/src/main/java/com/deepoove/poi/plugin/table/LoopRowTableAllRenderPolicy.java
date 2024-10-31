package com.deepoove.poi.plugin.table;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.policy.RenderPolicy;
import com.deepoove.poi.template.ElementTemplate;

import java.util.Map;

public class LoopRowTableAllRenderPolicy implements RenderPolicy {

    private String prefix;
    private String suffix;
    private boolean onSameLine;
    private boolean isSaveNextLine;

    public LoopRowTableAllRenderPolicy() {
        this(false);
    }

    public LoopRowTableAllRenderPolicy(boolean onSameLine) {
        this("[", "]", onSameLine);
    }

    public LoopRowTableAllRenderPolicy(boolean onSameLine, boolean isSaveNextLine) {
        this(onSameLine, isSaveNextLine, "[", "]");
    }

    public LoopRowTableAllRenderPolicy(String prefix, String suffix) {
        this(false, false, prefix, suffix);
    }


    public LoopRowTableAllRenderPolicy(String prefix, String suffix, boolean onSameLine) {
        this(onSameLine, false, prefix, suffix);
    }

    public LoopRowTableAllRenderPolicy(boolean onSameLine, boolean isSaveNextLine, String prefix, String suffix) {
        this.prefix = prefix;
        this.suffix = suffix;
        this.onSameLine = onSameLine;
        this.isSaveNextLine = isSaveNextLine;
    }


    @Override
    public void render(ElementTemplate eleTemplate, Object data, XWPFTemplate template) {
        int rendermode = 0;
        try {
            Map<String, Object> globalEnv = template.getEnvModel().getEnv();
            Object r = globalEnv.get(eleTemplate.getTagName() + "_rendermode");
            rendermode = r != null ? Integer.parseInt(r.toString()) : rendermode;
        } catch (NumberFormatException ignore) {
        }
        switch (rendermode) {
            case 1:
                new LoopExistedRowTableRenderPolicy(this.prefix, this.suffix, this.onSameLine, this.isSaveNextLine).render(eleTemplate, data, template);
                break;
            case 2:
                new LoopExistedAndFillRowTableRenderPolicy(this.prefix, this.suffix, this.onSameLine, this.isSaveNextLine).render(eleTemplate, data, template);
                break;
            case 3:
                new LoopRowTableAndFillRenderPolicy(this.prefix, this.suffix, this.onSameLine).render(eleTemplate, data, template);
                break;
            case 4:
                new LoopFullTableInsertFillRenderPolicy(this.prefix, this.suffix, this.onSameLine).render(eleTemplate, data, template);
                break;
            case 5:
                new LoopIncludeSubTableRenderPolicy(this.prefix, this.suffix, this.onSameLine).render(eleTemplate, data, template);
                break;
            case 6:
                new LoopCopyHeaderRowRenderPolicy(this.prefix, this.suffix, this.onSameLine).render(eleTemplate, data, template);
                break;
            default:
                new LoopRowTableRenderPolicy(this.prefix, this.suffix, this.onSameLine).render(eleTemplate, data, template);
        }
    }

    public String getPrefix() {
        return prefix;
    }

    public void setPrefix(String prefix) {
        this.prefix = prefix;
    }

    public String getSuffix() {
        return suffix;
    }

    public void setSuffix(String suffix) {
        this.suffix = suffix;
    }

    public boolean isOnSameLine() {
        return onSameLine;
    }

    public void setOnSameLine(boolean onSameLine) {
        this.onSameLine = onSameLine;
    }

    public boolean isSaveNextLine() {
        return isSaveNextLine;
    }

    public void setSaveNextLine(boolean saveNextLine) {
        isSaveNextLine = saveNextLine;
    }
}
