package com.deepoove.poi.plugin.table;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.policy.RenderPolicy;
import com.deepoove.poi.template.ElementTemplate;

import java.util.Map;

public class LoopRowTableAllRenderPolicy extends AbstractLoopRowTableRenderPolicy implements RenderPolicy {

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
        AbstractLoopRowTableRenderPolicy policy;
        switch (rendermode) {
            case 1:
                policy = new MultipleRowTableRenderPolicy();
                break;
            case 2:
                policy = new LoopExistedAndFillRowTableRenderPolicy(this);
                break;
            case 3:
                policy = new LoopRowTableAndFillRenderPolicy(this);
                break;
            case 4:
                policy = new LoopFullTableInsertFillRenderPolicy(this);
                break;
            case 5:
                policy = new LoopFullTableIncludeSubRenderPolicy(this);
                break;
            case 6:
                policy = new LoopCopyHeaderRowRenderPolicy(this);
                break;
            case 7:
                policy = new LoopCopyHeaderMutilpleRowRenderPolicy(this);
                break;
            case 8:
                policy = new LoopCopyHeaderMutilpleRowRenderSaveSuffixPolicy(this);
                break;
            default:
                policy = new LoopRowTableRenderPolicy(this);
        }
        policy.render(eleTemplate, data, template);
    }
}
