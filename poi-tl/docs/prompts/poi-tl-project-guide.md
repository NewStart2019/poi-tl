# poi-tl 项目开发提示词（AI 辅助开发指南）

## 一、项目概述

### 1.1 项目定位
poi-tl（poi template language）是基于 Apache POI 的 Word 模板引擎，使用 Word 模板 + 数据模型动态生成 docx 文档。
- 核心理念：TDO 模式 = Template（模板） + Data-model（数据） + Output（输出）
- 设计原则：所见即所得，模板样式完全保留
- 版本：1.13.8（自定义扩展版）

### 1.2 技术栈
| 技术 | 版本 | 用途 |
|---|---|---|
| Java | 1.8+ | 开发语言，源码级别兼容 JDK8 |
| Apache POI | 5.4.1 | Word 文档底层操作（poi-ooxml, poi-ooxml-full） |
| Maven | - | 项目构建工具 |
| SLF4J | 1.7.32 | 日志门面 |
| commons-lang3 | 3.12.0 | 工具类库 |
| Spring Expression | 5.3.26 | SpEL 表达式支持（provided 可选） |
| XMLBeans | - | POI 底层 XML 操作 |
| Batik | 1.17 | SVG 图片转换支持 |
| Jackson | 2.18.0 | JSON 处理 |
| JUnit 5 | 5.6.0 | 单元测试 |

### 1.3 官方文档
- 主站：https://deepoove.com/poi-tl/
- 新增功能文档：语雀内部文档

---

## 二、项目架构与模块划分

### 2.1 包结构说明
```
com.deepoove.poi
├── XWPFTemplate.java          # 核心门面类，用户入口
├── config/                    # 配置模块
│   ├── Configure.java         # 配置类
│   ├── ConfigureBuilder.java  # 构建器
│   ├── GramerSymbol.java      # 语法符号定义
│   └── PreRenderDataCastor.java
├── converter/                 # 数据转换器
├── data/                      # 数据模型（RenderData 体系）
│   ├── style/                 # 样式数据模型
│   ├── Texts.java             # 文本工厂
│   ├── Pictures.java          # 图片工厂
│   ├── Tables.java            # 表格工厂
│   ├── Rows.java              # 行工厂
│   ├── Cells.java             # 单元格工厂
│   └── *RenderData.java       # 各种渲染数据
├── exception/                 # 自定义异常
├── expression/                # 表达式引擎
├── plugin/                    # 插件体系
│   ├── table/                 # 表格相关插件（核心扩展区）
│   ├── bookmark/
│   ├── comment/
│   ├── field/
│   ├── pagination/
│   └── toc/
├── policy/                    # 渲染策略（策略模式）
│   ├── reference/             # 引用类图表策略
│   └── *RenderPolicy.java
├── render/                    # 渲染引擎
│   ├── compute/               # 数据计算
│   └── processor/             # 模板处理器
├── resolver/                  # 模板解析器
├── template/                  # 模板元素模型
├── util/                      # 工具类
│   ├── word/
│   │   └── WordTableUtils.java  # Word表格操作工具（核心扩展类，1500+行）
│   ├── TableTools.java        # 表格工具（官方原版）
│   ├── ParagraphUtils.java
│   ├── StyleUtils.java
│   ├── ReflectionUtils.java
│   ├── Preconditions.java     # 前置校验
│   ├── PoitlIOUtils.java
│   └── ...
└── xwpf/                      # POI XWPF 增强封装
    ├── NiceXWPFDocument.java  # 增强文档类
    ├── BodyContainer.java     # 正文容器抽象
    ├── XWPFParagraphWrapper.java
    ├── XWPFRunWrapper.java
    ├── XWPFTableRowWrapper.java
    └── ...
```

### 2.2 核心设计模式
1. **门面模式**：`XWPFTemplate` 作为统一入口，屏蔽内部复杂度
2. **策略模式**：`RenderPolicy` 接口，不同标签类型对应不同渲染策略
3. **工厂模式**：`Texts`, `Pictures`, `Tables`, `Rows`, `Cells` 等构建数据模型
4. **建造者模式**：`ConfigureBuilder`, `Style.builder()` 等
5. **模板方法模式**：`AbstractRenderPolicy`, `AbstractLoopRowTableRenderPolicy`
6. **插件机制**：可自定义 `RenderPolicy` 绑定到标签

---

## 三、编码规范与约定

### 3.1 通用规范
- **源码编码**：UTF-8
- **JDK 兼容性**：严格兼容 Java 8，不使用 Java 9+ 特性
- **License 头**：新增文件需添加 Apache License 2.0 版权头
- **日志**：统一使用 SLF4J 的 `LoggerFactory.getLogger()`
- **异常**：业务异常封装为 `RenderException` / `ResolverException`
- **空安全**：入参需考虑 null 情况，使用 `Preconditions` 做前置校验

### 3.2 命名规范
- 类名：大驼峰（PascalCase）
- 方法名：小驼峰（camelCase）
- 常量：全大写下划线分隔（UPPER_SNAKE_CASE）
- 布尔方法前缀：`isXXX`, `hasXXX`
- 查询方法前缀：`findXXX`, `getXXX`
- 修改方法前缀：`setXXX`, `addXXX`, `removeXXX`, `cleanXXX`, `copyXXX`
- 工具类：以 Utils / Tools 结尾

### 3.3 工具类编写规范（WordTableUtils 风格）
1. 工具类使用 `public class` + 私有构造（或默认构造）
2. 所有方法为 `public static`
3. 方法按功能分组：copy → clean → remove → find → set → merge
4. 重载方法提供便捷版本，最终调用最完整参数版本
5. 入参为 null 时打印 warn 日志并安全返回，不抛出 NPE
6. 操作底层 XML 时使用 `CT*` 类（如 CTRow, CTTc, CTTblPr 等）
7. 涉及反射使用 `ReflectionUtils.getValue()`

### 3.4 数据模型规范
- 所有渲染数据实现 `RenderData` 接口
- 使用 Builder 模式构建，提供工厂类（如 `Texts.of().create()`）
- 样式类统一放在 `data/style/` 包下

### 3.5 注释规范
- 公共方法必须有 Javadoc 注释，说明用途、参数、返回值
- 复杂逻辑添加行内注释
- TODO 标记使用 `// TODO 描述` 格式

---

## 四、核心标签与语法

### 4.1 默认标签语法
- 前后缀：`{{` 和 `}}`
- 文本标签：`{{var}}`
- 图片标签：`{{@var}}`
- 表格标签：`{{#var}}`
- 列表标签：`{{*var}}`
- 区块对（条件/循环）：`{{?var}} ... {{/var}}`
- 嵌套模板：`{{+var}}`

### 4.2 表格内动态行标签（自定义扩展）
- 前缀：`[`，后缀：`]`
- 占位符：`[fieldName]`
- 索引变量：`[_index]`（行号，从 1 开始）
- 表达式支持：`[a+b]`（SpEL 表达式）

### 4.3 动态表格渲染策略（rendermode）
通过 `xxx_subRecords_rendermode` 字段选择策略：
| rendermode | 策略类 | 说明 |
|---|---|---|
| 0（默认） | LoopRowTableRenderPolicy | 基础循环行渲染 |
| 1 | MultipleRowTableRenderPolicy | 多行模板插入行渲染 |
| 2 | LoopExistedAndFillRowTableRenderPolicy | 渲染已有行并填充空白页 |
| 3 | LoopRowTableAndFillRenderPolicy | 插入行并填充空白页 |
| 4 | LoopFullTableInsertFillRenderPolicy | 循环整表多行模板行渲染 |
| 5 | LoopFullTableIncludeSubRenderPolicy | 循环整表含子表数据渲染 |
| 6 | LoopCopyHeaderRowRenderPolicy | 循环表头行填充（支持 vmerge 跨列） |
| 7 | LoopCopyHeaderMutilpleRowRenderPolicy | 循环表头多行模板渲染 |
| 8 | LoopCopyHeaderMutilpleRowRenderSaveSuffixPolicy | 循环表头多行控制表尾填充 |

### 4.4 空白填充模式（mode）
通过 `xxx_subRecords_mode` 字段控制：
- 1：填充空白行（默认）
- 2：填充斜线
- 3：写文字"以下空白"在一行中间
- 4：写文字"以下空白"在最宽一列
- 5：自定义写入列（配合 `xxx_subRecords_write_col` 指定列号）

---

## 五、WordTableUtils 工具类使用指南

### 5.1 功能分组
```
copy（复制）    → copyTable, copyLine, copyCell, copyParagraph, copyRun
clean（清理）   → cleanRowTextContent, cleanCellContent, cleanParagraphContent
remove（删除）  → removeTable, removeRow, removeLastRow, removeParagraph, removeRun
find（查找）    → findRowIndex, findRowHeight, findCellWidth, findVerticalMergedRows
set（设置）     → setTableRow, setTableRowHeight, setCellWidth, setCellVmerge, setDiagonalBorder
merge（合并）   → mergeCellsHorizontal, mergeCellsVertically, mergeMutipleLine
```

### 5.2 关键方法说明
- `copyTable(doc, sourceTable, isTail)`：复制整个表格到文档末尾
- `removeRow(table, rowIndex)`：删除指定行
- `setTableRow(table, row, pos)`：设置指定位置的行对象（替换）
- `findRowIndex(row)`：获取行在表格中的索引
- `mergeCellsVertically(table, col, fromRow, toRow)`：纵向合并单元格
- `setDiagonalBorder(cell)`：设置单元格斜线边框

### 5.3 行操作底层原理
- 行列表存储在 `tableRows` 私有字段，通过反射访问
- XML 层面操作 `CTTbl.setTrArray(pos, ctRow)`
- 移动行需要同时更新 Java 对象列表和底层 XML 结构

---

## 六、新增/修改代码 Checklist

### 6.1 新增工具方法
- [ ] 方法为 `public static`
- [ ] 入参 null 校验 + warn 日志
- [ ] 完整 Javadoc 注释
- [ ] 重载方法提供便捷版本
- [ ] 异常情况安全返回，不中断主流程

### 6.2 新增渲染策略
- [ ] 继承 `AbstractLoopRowTableRenderPolicy`
- [ ] 实现 `render()` 方法
- [ ] 在 `LoopRowTableAllRenderPolicy` 的 switch 中注册
- [ ] 添加对应测试用例

### 6.3 代码安全与健壮性
- [ ] 数组/列表越界检查
- [ ] 空指针防护
- [ ] 资源关闭（流、cursor 等）
- [ ] 索引参数合法性校验
- [ ] 边界条件处理（首行、末行、空表）

---

## 七、测试规范
- 测试类放在 `src/test/java/com/deepoove/poi/tl/` 对应包下
- 使用 JUnit 5（`@Test` from junit-jupiter-api）
- 测试输出文件统一放在 `target/` 目录
- 测试模板资源放在 `src/test/resources/`

---

## 八、自定义扩展历史变更（2.3.x 系列）

| 日期 | 变更内容 |
|---|---|
| 2024-10-23 | 渲染策略选择、行号从1开始 |
| 2024-10-25 | 表头字段id去掉后缀下划线 |
| 2024-10-31 | 多行表头、删除行策略 |
| 2024-11-08 | 解决尾部空白行导致空白页 |
| 2024-11-11 | SpEL函数、子表数据覆盖修复 |
| 2024-11-13 | 复制多行渲染 + 空白填充模式 |
| 2024-11-19 | 多行渲染边框控制 |
| 2024-11-21 | 模式6-7优化、新增模式8 |
| 2024-11-23 | 策略1改为多行模板渲染 |
| 2024-11-25 | 所有策略重写，解决插入位置错乱 |
| 2024-12-04 | 填充模式4（最宽列写"以下空白"） |
| 2024-12-18 | 策略6自定义合并列标记 |
| 2025-02-11 | 策略6 vmerge 跨列标签 |
| 2025-05-07 | 表头跨列统计修复、字体修复 |
| 2025-07-25 | 删除整个表格策略 |
| 2026-03-12 | 删除行补行功能 |
| 2026-07-07 | 填充模式非法值重置为1 |
| 2026-07-15 | 自定义"以下空白"列、历史 bug 修复 |

---

## 九、常用代码片段参考

### 9.1 获取表格行索引
```java
int rowIndex = WordTableUtils.findRowIndex(row);
```

### 9.2 删除表格行
```java
WordTableUtils.removeRow(table, rowIndex);
```

### 9.3 设置行高
```java
WordTableUtils.setTableRowHeight(row, UnitUtils.point2Twips(24), STHeightRule.EXACT);
```

### 9.4 复制行内容
```java
WordTableUtils.copyLineContent(sourceRow, targetRow, templateRowIndex);
```

### 9.5 纵向合并单元格
```java
WordTableUtils.mergeCellsVertically(table, colIndex, startRow, endRow);
```

---

## 十、开发注意事项

1. **POI 版本兼容性**：代码需兼容 POI 5.x API，不使用已废弃方法
2. **XML 操作谨慎**：直接操作 CT* 类时注意 XML 结构正确性
3. **资源释放**：XmlCursor 使用后必须 `close()`，文档流必须关闭
4. **反射使用**：访问 POI 私有字段用 `ReflectionUtils`，注意字段名版本差异
5. **空表处理**：操作表格前判断行数/列数，避免 IndexOutOfBounds
6. **跨文档复制**：复制元素到不同 XWPFDocument 时注意引用关系
7. **样式保留**：复制操作尽量保留原始样式（isIncludeStyle 参数）
8. **性能考虑**：大表格循环操作注意性能，避免频繁 XML 解析
