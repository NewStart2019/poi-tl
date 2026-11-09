package com.deepoove.poi.plugin.table;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.exception.RenderException;
import com.deepoove.poi.policy.RenderPolicy;
import com.deepoove.poi.render.compute.EnvModel;
import com.deepoove.poi.render.compute.SpELRenderDataCompute;
import com.deepoove.poi.template.BlockTemplate;
import com.deepoove.poi.template.ElementTemplate;
import com.deepoove.poi.template.MetaTemplate;
import com.deepoove.poi.template.run.RunTemplate;
import com.deepoove.poi.util.TableTools;
import com.deepoove.poi.util.word.WordTableUtils;
import org.apache.poi.xwpf.usermodel.TableRowHeightRule;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHeight;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTrPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STHeightRule;

import java.math.BigInteger;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.IdentityHashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Set;

/**
 * 表格行删除策略。
 *
 * <p>占位符规则：
 * <ul>
 *     <li>占位符取值为 {@code null}：删除所在行；</li>
 *     <li>占位符取值为构造时传入的默认字符串（如 "——"）：删除所在行；</li>
 *     <li>表达式结果为 {@link Boolean#FALSE}：删除所在行；</li>
 *     <li>其他取值（含 {@link Boolean#TRUE}）：保留所在行。</li>
 * </ul>
 * 仅适用于表格，且占位符需位于单独的段落中。
 *
 * <p>补行规则：当数据中存在 {@code open_addRow} 时，每删除一行会按行高补充空白行：
 * 插入行数 = {@code Max(删除行高 / 复制行高, 1)}（整数除法向下取整），每行高度均等于复制行高度；
 * 若删除行与复制行都没有固定行高，则兜底插入一行、高度与复制行一致。
 * 默认在表格尾部补充，若存在 {@code xxxx_insertPosition} 字段，则在删除行之后指定位置补充。
 *
 * <p><b>本版本修复要点：</b>
 * <ol>
 *     <li>行索引以底层 {@link CTRow} 的对象身份在 {@code CTTbl.getTrArray()} 中匹配，
 *     不再依赖 {@code table.getRows().indexOf(包装行)}，避免结构编辑后包装对象与底层 XML 脱节而删错行；</li>
 *     <li>默认字符串比较使用 {@link Objects#equals}，不再使用 {@code ==} 引用相等；</li>
 *     <li>补行按行高计算插入行数（Max(删除行高/复制行高,1)，无固定行高兜底 1 行），
 *     采用"深拷贝结构 + 清空文本"的方式生成真正的空白行，不替换、不清理已有行包装，
 *     从根源上避免 RunTemplate 持有的行引用被孤立；</li>
 *     <li>空白行会移除继承到的 vMerge 标记，避免其所属合并组被删除后留下悬空 continue、
 *     被相邻合并组错误吸收而在内容区出现"延伸"的空白行；</li>
 *     <li>垂直合并（vMerge）分组借助全局删除集合判断：整组都删除时不再迁移 restart 标签，
 *     仅当组内存在保留行时才把 restart 内容提升到第一个保留行；</li>
 *     <li>不再改写全局 {@code config} 的计算工厂，消除对共享配置的副作用。</li>
 * </ol>
 */
public class RemoveTableRowRenderPolicy implements RenderPolicy {

    /** 本策略绑定的占位符前缀符号。 */
    private static final Character POLICY_SIGN = '$';
    /** 是否开启补行的开关字段名。 */
    private static final String ADD_ROW_FLAG = "open_addRow";
    /** 指定补行位置字段的后缀。 */
    private static final String INSERT_POSITION_SUFFIX = "_insertPosition";

    private final String defaultDeleteValue;

    /**
     * 每个模板对应的"待删除行"集合（以底层 CTRow 为身份）。
     * 用 IdentityHashMap 保证按对象身份存取，且不长期持有无关文档。
     */
    private final Map<XWPFTemplate, Set<CTRow>> plannedDeletes = new IdentityHashMap<>();

    public RemoveTableRowRenderPolicy() {
        this(null);
    }

    public RemoveTableRowRenderPolicy(String defaultDeleteValue) {
        this.defaultDeleteValue = defaultDeleteValue;
    }

    @Override
    public void render(ElementTemplate eleTemplate, Object data, XWPFTemplate template) {
        RunTemplate runTemplate = (RunTemplate) eleTemplate;
        XWPFRun run = runTemplate.getRun();
        try {
            if (!TableTools.isInsideTable(run)) {
                throw new IllegalStateException(
                    "The template tag " + runTemplate.getSource() + " must be inside a table");
            }

            Map<String, Object> globalEnv = template.getEnvModel().getEnv();

            // 首次执行时，对象关系尚未被结构编辑破坏，一次性规划出所有待删除行，
            // 得到全局删除集合，供 vMerge 跨行分组判断"是否整组删除"。
            Set<CTRow> rowsToDelete = plannedDeletes.get(template);
            if (rowsToDelete == null) {
                rowsToDelete = planDeletions(template, globalEnv);
                plannedDeletes.put(template, rowsToDelete);
            }

            // 独立创建 SpEL 计算器，不改写全局 config。
            SpELRenderDataCompute dataCompute = new SpELRenderDataCompute(
                EnvModel.of(null, globalEnv), false);
            Object value = dataCompute.compute(runTemplate.getTagName());

            XWPFParagraph paragraph = (XWPFParagraph) run.getParent();
            XWPFTableCell tagCell = (XWPFTableCell) paragraph.getBody();
            XWPFTableRow tableRow = tagCell.getTableRow();
            XWPFTable table = tableRow.getTable();

            if (!shouldDelete(value)) {
                // 保留行：仅清空占位符文本并移除其所在段落，保留整行。
                run.setText("", 0);
                tagCell.removeParagraph(tagCell.getParagraphs().indexOf(paragraph));
                return;
            }

            // 删除行：以底层 CTRow 身份定位真实当前索引，避免包装对象失配。
            int rowIndex = WordTableUtils.findRowIndex(table, tableRow.getCtRow());
            if (rowIndex < 0) {
                // 兜底：该行已被移除（例如多个标签共用一行），直接跳过。
                return;
            }

            // 在删除前捕获删除行的高度信息（删除后该行即从表中移除，无法再读取）。
            int deletedHeight = tableRow.getHeight();
            boolean deletedFixed = hasFixedHeight(tableRow);

            deleteRow(table, rowIndex, rowsToDelete);

            // 补行逻辑。
            if (globalEnv.get(ADD_ROW_FLAG) != null) {
                Object positionCfg = globalEnv.get(runTemplate.getTagName() + INSERT_POSITION_SUFFIX);
                int insertPosition;
                if (positionCfg == null) {
                    // 默认在表格尾部补充空白行。
                    insertPosition = table.getNumberOfRows();
                } else {
                    // 在删除位置之后第 position 行插入空白行。
                    insertPosition = rowIndex + ((Number) positionCfg).intValue() - 1;
                }
                // 复制行：插入位置当前所在行（尾部插入时为最后一行），提供结构与目标高度。
                int copyIndex = Math.min(Math.max(insertPosition, 0), table.getNumberOfRows() - 1);
                if (copyIndex < 0) {
                    // 表已空，无可用复制行，跳过补行。
                    return;
                }
                XWPFTableRow copyRow = table.getRow(copyIndex);
                insertBlankRows(table, insertPosition, deletedHeight, deletedFixed, copyRow);
            }
        } catch (Exception e) {
            throw new RenderException("Remove line failure: " + e.getMessage(), e);
        }
    }

    /**
     * 判断占位符取值是否表示删除行。
     *
     * @param value 占位符计算结果
     * @return true 表示删除；false 表示保留
     */
    private boolean shouldDelete(Object value) {
        if (value == null) {
            return true;
        }
        if (value instanceof Boolean) {
            // 表达式结果为 false 删除，为 true 保留。
            return Boolean.FALSE.equals(value);
        }
        // 使用值相等而非引用相等，兼容 JSON、new String 等非驻留字符串。
        return Objects.equals(value, defaultDeleteValue);
    }

    /**
     * 在任何结构编辑之前，基于当前一致的对象关系计算全部待删除行。
     *
     * @param template  当前模板
     * @param globalEnv 全局数据
     * @return 待删除行的底层 CTRow 集合
     */
    private Set<CTRow> planDeletions(XWPFTemplate template, Map<String, Object> globalEnv) {
        SpELRenderDataCompute dataCompute = new SpELRenderDataCompute(
            EnvModel.of(null, globalEnv), false);
        Set<CTRow> result = new HashSet<>();
        for (RunTemplate rt : collectRemoveRunTemplates(template.getElementTemplates())) {
            if (shouldDelete(dataCompute.compute(rt.getTagName()))) {
                XWPFRun r = rt.getRun();
                XWPFTableCell cell = (XWPFTableCell) ((XWPFParagraph) r.getParent()).getBody();
                result.add(cell.getTableRow().getCtRow());
            }
        }
        return result;
    }

    /**
     * 递归收集所有绑定本策略（符号为 $）的 RunTemplate，兼容 Iterable/Block 嵌套结构。
     */
    private List<RunTemplate> collectRemoveRunTemplates(List<MetaTemplate> templates) {
        List<RunTemplate> result = new ArrayList<>();
        if (templates == null) {
            return result;
        }
        for (MetaTemplate mt : templates) {
            if (mt instanceof RunTemplate) {
                RunTemplate rt = (RunTemplate) mt;
                if (POLICY_SIGN.equals(rt.getSign())) {
                    result.add(rt);
                }
            } else if (mt instanceof BlockTemplate) {
                // IterableTemplate 亦为 BlockTemplate 的子类，在此统一递归处理。
                result.addAll(collectRemoveRunTemplates(((BlockTemplate) mt).getTemplates()));
            }
        }
        return result;
    }


    /**
     * 删除指定行，处理垂直合并（vMerge）的 restart 单元格。
     *
     * @param table        表格
     * @param rowIndex     待删除行索引
     * @param rowsToDelete 全局待删除行集合
     */
    private void deleteRow(XWPFTable table, int rowIndex, Set<CTRow> rowsToDelete) {
        XWPFTableRow row = table.getRow(rowIndex);
        int cellCount = row.getTableCells().size();
        for (int col = cellCount - 1; col >= 0; col--) {
            XWPFTableCell cell = row.getCell(col);
            if (WordTableUtils.findVMerge(cell) != 2) {
                continue;
            }
            // 当前单元格为 vMerge restart，查找合并组内第一个保留行。
            int firstSurviving = firstSurvivingContinuation(table, rowIndex, col, rowsToDelete);
            if (firstSurviving >= 0) {
                // 组内有保留行：把 restart 内容与属性提升到第一个保留行，保证合并语义延续。
                WordTableUtils.moveRestartCell(table, rowIndex, col, firstSurviving);
            }
            // 整组都删除：不迁移任何内容，随整行一并删除即可。
        }
        table.removeRow(rowIndex);
    }

    /**
     * 沿 vMerge continue 向下查找合并组内第一个被保留的行。
     *
     * <p>该方法直接读取单元格的 vMerge 属性，并对空行、列错位等情况做防御处理，
     * 不依赖会在空行上抛出越界异常的 {@code findVerticalMergedRows}。
     *
     * @param table        表格
     * @param rowIndex     restart 所在行索引
     * @param col          restart 所在列索引
     * @param rowsToDelete 全局待删除行集合
     * @return 第一个保留行索引；合并组全部删除时返回 -1
     */
    private int firstSurvivingContinuation(XWPFTable table, int rowIndex, int col,
                                           Set<CTRow> rowsToDelete) {
        int total = table.getNumberOfRows();
        for (int k = rowIndex + 1; k < total; k++) {
            XWPFTableRow r = table.getRow(k);
            List<XWPFTableCell> cells = r.getTableCells();
            // 容忍空行或列错位：取不到对应列即视为合并组结束。
            if (col >= cells.size()) {
                break;
            }
            XWPFTableCell c = cells.get(col);
            // 仅 continue（值为 1）属于当前合并组；遇到 restart 或无 vMerge 即结束。
            if (WordTableUtils.findVMerge(c) != 1) {
                break;
            }
            if (!rowsToDelete.contains(r.getCtRow())) {
                return k;
            }
        }
        return -1;
    }

    /**
     * 按行高计算并插入空白行。
     *
     * <p>插入行数 = {@code Max(删除行高 / 复制行高, 1)}（整数除法，向下取整），
     * 每行高度均设置为复制行高度（固定行高 EXACT）。
     * 若删除行与复制行都没有设置固定行高，则兜底插入一行，结构与高度沿用复制行，不强制固定。
     *
     * <p>空白行均深拷贝复制行结构、清空文本并剥离 vMerge，不替换、不清理任何已有行，
     * 避免孤立已解析的 RunTemplate。
     *
     * @param table         表格
     * @param position      期望插入位置
     * @param deletedHeight 删除行高度（twips）
     * @param deletedFixed  删除行是否为固定行高
     * @param copyRow       复制行（提供结构与目标高度）
     */
    private void insertBlankRows(XWPFTable table, int position, int deletedHeight,
                                 boolean deletedFixed, XWPFTableRow copyRow) {
        if (copyRow == null || table.getNumberOfRows() == 0) {
            return;
        }
        int copyHeight = copyRow.getHeight();
        boolean copyFixed = hasFixedHeight(copyRow);

        int count;
        boolean forceFixedHeight;
        if (!deletedFixed && !copyFixed) {
            // 删除行与复制行都没有固定行高：兜底插入一行，高度与复制行一致（沿用其结构，不强制固定）。
            count = 1;
            forceFixedHeight = false;
        } else {
            // 插入行数 = Max(删除行高 / 复制行高, 1)，整数除法向下取整。
            int ratio = (copyHeight > 0) ? (deletedHeight / copyHeight) : 1;
            count = Math.max(ratio, 1);
            forceFixedHeight = true;
        }

        // 越界兜底：插入位置限制在 [0, 行数]。
        int insertPos = Math.min(Math.max(position, 0), table.getNumberOfRows());

        for (int i = 0; i < count; i++) {
            XWPFTableRow blank = table.insertNewTableRow(insertPos);
            CTRow copy = (CTRow) copyRow.getCtRow().copy();
            WordTableUtils.cleanRowTextKeepStructure(copy);
            // 空白行不继承任何垂直合并标记，保证其为独立的一行。
            stripVerticalMerge(copy);
            if (forceFixedHeight && copyHeight > 0) {
                // 每行高度 = 复制行高度，固定行高（EXACT）。
                applyFixedRowHeight(copy, copyHeight);
            }
            blank.getCtRow().set(copy);
            // 下一行紧随其后插入，保证多行连续且位置不重叠。
            insertPos++;
        }
    }

    /**
     * 判断行是否设置了固定行高（高度规则为 EXACT）。
     *
     * @param row 待判断的行
     * @return true 表示固定行高；AT_LEAST 与 AUTO 均不视为固定行高
     */
    private boolean hasFixedHeight(XWPFTableRow row) {
        return row.getHeightRule() == TableRowHeightRule.EXACT;
    }

    /**
     * 在底层行上设置固定行高（EXACT）。
     *
     * @param ctRow       目标底层行
     * @param heightTwips 高度（twips）
     */
    private void applyFixedRowHeight(CTRow ctRow, int heightTwips) {
        CTTrPr trPr = ctRow.isSetTrPr() ? ctRow.getTrPr() : ctRow.addNewTrPr();
        CTHeight height = trPr.sizeOfTrHeightArray() == 0
            ? trPr.addNewTrHeight() : trPr.getTrHeightArray(0);
        height.setVal(BigInteger.valueOf(heightTwips));
        height.setHRule(STHeightRule.EXACT);
    }

    /**
     * 移除行内所有单元格的垂直合并（vMerge）标记，保留 gridSpan、宽度等其他单元格属性。
     *
     * @param row 待处理的底层行（通常为深拷贝出的空白行副本）
     */
    private void stripVerticalMerge(CTRow row) {
        for (CTTc tc : row.getTcList()) {
            if (tc.isSetTcPr() && tc.getTcPr().isSetVMerge()) {
                tc.getTcPr().unsetVMerge();
            }
        }
    }
}
