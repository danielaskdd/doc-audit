# Word 文档内容提取与智能分块指南

本文面向需要将 Word 文档交给 LLM 审核的读者，解释**内容提取需要解决的问题**、**智能分块的优化策略**，以及 `_blocks.jsonl` 输出的**字段含义与典型样例**。

> 本文内容基于 `skills/doc-audit/scripts/parse_document.py` 的现有实现整理。

---

## 1. 概述：为什么需要“内容提取 + 智能分块”

在文档审核场景中，Word 文档有三个天然难点：

- **结构复杂**：标题层级、自动编号、合并单元格、跨页表头并存。
- **内容异构**：正文、表格、图片、公式、上下标共存。
- **上下文限制**：LLM 有输入长度上限，超长内容必须拆分。

因此，从 `.docx` 到 `_blocks.jsonl` 的过程需要同时解决：

1. **结构还原**：识别真实标题层级与编号语义。
2. **内容保真**：表格、图片、公式、上下标等信息不丢失。
3. **可审计分块**：既覆盖完整上下文，又不超出模型窗口。

最终输出的 `_blocks.jsonl` 是审计系统的基础输入。

---

## 2. 内容提取需要解决哪些问题

### 2.1 标题层级识别

- 使用 `outlineLvl` 判断标题层级（而不是样式名字符串）。
- 支持样式继承（`basedOn`）以补全层级信息。
- `outlineLvl >= 9` 视为正文。

**效果**：不同层级标题会形成后续分块的语义边界。

### 2.2 自动编号还原

- 通过 `NumberingResolver` 还原 Word 自动编号（章节号、列表号、图表号）。
- 让审核内容更接近用户的视觉阅读结果。

**效果**：例如 `2.1 Penalty Clause` 会保留“2.1”。

### 2.3 表格完整提取

- 支持纵向/横向合并单元格，尽量保持语义，便于 LLM 理解。
- 支持跨页表头（`w:tblHeader`）。表格分片后仍可保留表头，帮助 LLM 理解上下文。
- 表格内容输出为 JSON 二维数组，并嵌入 `<table>...</table>`，以 LLM 更容易理解的方式提供结构化内容。

**效果**：表格语义可被较完整地保留，适合审计规则匹配。

### 2.4 特殊内容保真

- **图片**：`<drawing id="..." name="..." path="..." format="..." />`
  - 内嵌图片（`r:embed`）：导出到 `<blocks_stem>.image/`，`path` 写相对路径
  - 外链图片（`r:link`）：不下载，`path` 写原始外链 URL
  - `format`：由 content-type/文件扩展名归一化得到，常见取值包括
    `png`、`jpeg`、`emf`、`wmf`、`gif`、`bmp`、`tiff`、`webp`、`svg`
  - 若无法判定格式，或该 drawing 不包含图片 blip（如图表形状），则省略 `format` 属性
- **公式**：OMML → LaTeX，包裹 `<equation>...</equation>`
- **上下标**：`<sup>...</sup>` / `<sub>...</sub>`

**效果**：减少“公式缺失/标记丢失”导致的误判。

### 2.5 段落定位锚点

每个段落依赖 Word 的 `w14:paraId`：

- 保证块的起止位置可定位（`uuid` / `uuid_end`）。
- 为后续审核结果标注大幅缩小检索范围，减少标注失败。
- 缺少 `paraId` 会直接报错（通常是非标准 Word 2013+ 文档，需要在 Word 中打开后另存）。
- 普通段落：取首段与尾段的 `paraId`。
- 表格：取首单元格第一段和末单元格最后一段的 `paraId`；对于垂直合并单元格，`paraId` 需要回溯到合并组的首个单元格，以便后续标注时获得正确编辑位置。

---

## 3. 智能分块策略（性能优化核心）

**为什么不能简单按标题切分？**

- 标题之间的正文可能过长，超出模型输入限制。
- 表格通常是“超大块”，不能在任意位置随意切分。
- 如果按标题完全割裂内容，粒度过细会导致审核阶段难以发现相邻段落间的勾稽关系问题。

**智能分块需要达到的目标**

- **充分利用 LLM 上下文窗口**：在不突破 `MAX_BLOCK_CONTENT_TOKENS` 的前提下，让块大小尽量接近 `IDEAL_BLOCK_CONTENT_TOKENS`，减少上下文割裂。
- **优先保持文档层级语义**：让 LLM 理解每段正文对应的标题及完整层级链；确保文本块以标题开始；确保块内不会混入更高层级标题所属正文。
- **避免表格与正文上下文断裂**：尽量让表格与前后正文保持关联。

因此采用四阶段流水线：

### 阶段 A：标题驱动的初始切块

- 标题段落触发块边界。
- 标题文本本身也进入正文（作为块首段）。
- 若标题之间没有正文，不落盘空块。

**目的**：保证块具有清晰语义边界。

**标题识别逻辑**：

- 主来源：标题段（`outlineLvl` 0-8）文本；自动编号先由 `NumberingResolver` 拼接到前缀。
- 非标题开头默认：`Preface/Uncategorized`。
- 长块拆分时，锚点段落会成为后续子块的新 `heading`。
- 大表分片的非首片会追加后缀：`[表格片段N]`。

> 为保证审核阶段可靠性，当前实现会在标题文字超过 200 个字符时报错退出（不会走“超长截断”流程）。

### 阶段 B：大表按行切片

触发条件：表格 token 估算值 `> 5000`。

策略：

- 仅按**行边界**切片。
- 首片与前文合并，中间片独立成块，末片与后文合并。
- 中间片/末片可携带 `table_header`（跨页表头行）。

**目的**：避免超长表格“撑爆上下文”。

> 非首片标题会在原标题后追加：`[表格片段<序号>]`。`table_header` 的内容来源于 Word 的跨页表头；如果未设置跨页表头，则不会添加 `table_header` 字段。

**表格分片后，“正文与表格是否同块”**：

- 未分片表格：与前后正文同块（若存在）。
- 分片表格首片：与表前正文同块（若存在）。
- 分片表格中间片：独立成块，不与正文同块。
- 分片表格末片：与表后正文同块（若存在）。

### 阶段 C：长块锚点拆分

触发条件：块 token 估算值 `> 8000`。

策略：

- 选用短段落（`<= 100` 字符，视为可充当小标题）作为锚点。
- 锚点既作为新块 `heading`，也保留在正文中。
- 不对表格做二次切分。

**目的**：在保持语义连贯的同时分散长度。段落拆分规模不必大于表格切片规模，以便尽量让表格与前后正文落在同一语义上下文内。

### 阶段 D：小块合并（保持层级语义）

策略：

- 自底向上，先合并同级小块，再允许跨级吸收。
- 合并后不能超过 8000 token。
- 遇到 `table_chunk_role="middle"` 时禁止合并。
- Tail Absorption（尾部吸收）：若当前块尾部存在同级“小尾巴”（`< 1000` token），尝试吸收，避免遗留无法继续合并的碎片。

**目的**：减少过碎分块，提升审核准确性。

**最终分块效果（总体）**：

- 块数量通常减少，单块上下文更完整。
- 大多数块会落在“接近 6000 token，且不超过 8000 token”的区间。
- 表格中间片段不受小块合并影响。

### 固定层级模式（`--fixlevel=N`）

若使用 `--fixlevel=N`：

- 仅标题层级 `<= N` 触发切块。
- 不执行表格分片与小块合并。
- 最后一块仍可能触发长块拆分。

该模式仅用于特殊文本分块场景，文档审核主流程不使用此模式。

---

## 4. 输出文件格式：`_blocks.jsonl`

文件采用 **JSON Lines**：

1. **第 1 行**：元数据（`type: "meta"`）
2. **第 2 行起**：文本块（`type: "text"`）

### 4.1 元数据行

```json
{
  "type": "meta",
  "source_file": "/absolute/path/to/document.docx",
  "source_hash": "sha256:f59469543f31d54f...",
  "parsed_at": "2026-01-28T18:17:32.175446"
}
```

| 字段 | 类型 | 说明 |
|------|------|------|
| `type` | string | 固定为 `"meta"` |
| `source_file` | string | 源文档绝对路径 |
| `source_hash` | string | 源 `.docx` 的 SHA256：`sha256:<hex>` |
| `parsed_at` | string | 本地时间戳（无时区） |

### 4.2 文本块行

```json
{
  "uuid": "682A7C9F",
  "uuid_end": "5A77B586",
  "heading": "2.1 Penalty Clause",
  "content": "If Party B delays payment...",
  "type": "text",
  "parent_headings": ["Chapter 2 Contract Terms"],
  "level": 2,
  "table_chunk_role": "none",
  "table_header": [["col-1", "col-2", "col-3"]]
}
```

| 字段 | 类型 | 说明 |
|------|------|------|
| `uuid` | string | 块首段落 `paraId` |
| `uuid_end` | string | 块尾段落 `paraId` |
| `heading` | string | 当前块标题（含自动编号） |
| `content` | string | 多段文本拼接后的内容 |
| `type` | string | 固定为 `"text"` |
| `parent_headings` | list | 祖先标题链 |
| `level` | int | 标题层级（1-9） |
| `table_chunk_role` | string | 表格分片角色：`none/first/middle/last` |
| `table_header` | list | 可选，表头行（跨页表头） |

---

## 5. 典型分块样例

### 5.1 普通文本块

```json
{
  "uuid": "AA12B3C4",
  "uuid_end": "DD56E7F8",
  "heading": "1.1 项目范围",
  "content": "本项目包含...\n\n本节说明...",
  "type": "text",
  "parent_headings": ["第一章 总则"],
  "level": 2,
  "table_chunk_role": "none"
}
```

### 5.2 含表格的文本块（未分片）

```json
{
  "uuid": "1A2B3C4D",
  "uuid_end": "5E6F7A8B",
  "heading": "2.2 技术指标",
  "content": "指标说明如下：\n<table>[['...'], ['...']]</table>",
  "type": "text",
  "parent_headings": ["第二章 技术要求"],
  "level": 2,
  "table_chunk_role": "none"
}
```

### 5.3 大表分片（中间片独立输出）

```json
{
  "uuid": "1234ABCD",
  "uuid_end": "5678EF90",
  "heading": "3.2 数据范围 [表格片段1]",
  "content": "<table>[['...'], ['...']]</table>",
  "type": "text",
  "parent_headings": ["第3章 数据要求"],
  "level": 2,
  "table_chunk_role": "middle",
  "table_header": [["列1", "列2", "列3"]]
}
```

### 5.4 长块锚点拆分后的子块

```json
{
  "uuid": "A1B2C3D4",
  "uuid_end": "E5F6A7B8",
  "heading": "背景",
  "content": "背景\n...",
  "type": "text",
  "parent_headings": ["3.1 项目描述"],
  "level": 3,
  "table_chunk_role": "none"
}
```

### 5.5 小块合并后的结果

```json
{
  "uuid": "AAAA1111",
  "uuid_end": "BBBB2222",
  "heading": "4.1 风险分析",
  "content": "4.1 风险分析...\n\n4.2 风险控制措施...\n\n...",
  "type": "text",
  "parent_headings": ["第四章 风险"],
  "level": 2,
  "table_chunk_role": "none"
}
```

---

## 6. 阈值参数速查表

| 常量 | 值 | 作用 |
|------|----|------|
| `MAX_HEADING_LENGTH` | 200 字符 | 标题长度上限 |
| `MAX_ANCHOR_CANDIDATE_LENGTH` | 100 字符 | 长块拆分锚点上限 |
| `IDEAL_BLOCK_CONTENT_TOKENS` | 6000 | 理想块大小 |
| `MAX_BLOCK_CONTENT_TOKENS` | 8000 | 块最大限制 |
| `SMALL_TAIL_THRESHOLD` | 1000 | 尾部吸收阈值（避免碎片） |
| `TABLE_IDEAL_TOKENS` | 3000 | 表格切片目标 |
| `TABLE_MAX_TOKENS` | 5000 | 表格切片触发阈值 |
| `TABLE_MIN_LAST_CHUNK_TOKENS` | 1600 | 末片最小阈值（避免碎片） |

---

## 附录：Token 估算方法

块大小基于 `estimate_tokens()` 的**启发式估算**，并非模型的精确 token 计数：

- 中文字符约 0.75 token/字
- JSON 结构字符约 1 token/字符
- 其他字符约 0.4 token/字符
- 额外增加约 5% buffer + 固定偏移

该方法可在不同模型间保持分块策略相对稳定、可控。
