# Word 内容提取结果（分块）文件说明

本文档基于 `skills/doc-audit/scripts/parse_document.py` 当前实现整理，描述 `_blocks.jsonl` 的字段含义与“智能分块”真实行为。

## 文件格式

输出为 JSONL（JSON Lines）：

1. 第 1 行：元数据（`type: "meta"`）
2. 第 2 行起：文本块（`type: "text"`）

---

## 1. 元数据行（Meta）

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
| `source_file` | string | 源文档绝对路径（`Path(file).resolve()`） |
| `source_hash` | string | 源 `.docx` 二进制文件 SHA256，格式 `sha256:<hex>` |
| `parsed_at` | string | `datetime.now().isoformat()`，本地时间字符串（无时区） |

备注：
- `run_audit.py --resume` 会使用 `source_hash` 防止“文档变了但继续旧结果”。

---

## 2. 文本块行（Text Block）

```json
{
  "uuid": "682A7C9F",
  "uuid_end": "5A77B586",
  "heading": "2.1 Penalty Clause",
  "content": "If Party B delays payment...",
  "type": "text",
  "parent_headings": ["Chapter 2 Contract Terms"],
  "level": 2,
  "table_chunk_role": "none"
}
```

当文本块来自“表格分片（非首片）”时，可能包含 `table_header`，例如：

```json
{
  "uuid": "1234ABCD",
  "uuid_end": "5678EF90",
  "heading": "3.2 数据范围 [表格片段1]",
  "content": "<table>[[\"...\"],[\"...\"]]</table>",
  "type": "text",
  "parent_headings": ["第3章 数据要求"],
  "level": 2,
  "table_chunk_role": "middle",
  "table_header": [["列1", "列2", "列3"]]
}
```

### 2.1 `uuid` / `uuid_end`

- 来源：Word 段落 `w14:paraId`。
- `uuid`：块首锚点；`uuid_end`：块尾锚点。
- 普通段落：取首段和尾段 paraId。
- 表格：`uuid` 用 `para_ids` 的首个有效值；`uuid_end` 用 `para_ids_end` 的最后有效值（单元格可能含多段）。
- 如果缺失 paraId（常见于非 Word 2013+ 生成文件），脚本直接报错退出。

### 2.2 `heading`

- 主来源：标题段（`outlineLvl` 0-8）文本，自动编号会先经 `NumberingResolver` 拼到前面。
- 非标题开头默认：`Preface/Uncategorized`。
- 长块切分时，锚点段落会成为后续分块的新 `heading`。
- 大表分片的非首片会追加后缀：`[表格片段N]`。

注意：
- 当前实现先 `validate_heading_length()`（超 200 字符即退出），再调用 `truncate_heading()`。因此“超长后截断继续”在主流程基本不会发生，实际行为是直接失败。

### 2.3 `content`

- 多段之间用 `\n` 拼接；章节之间使用 `\n\n`拼接。
- 特殊内容的内嵌标记：
  - 表格：`<table>[[...]]</table>`
  - 图片：`<drawing id="..." name="..." />`
  - 公式：`<equation>...</equation>`
  - 上下标：`<sup>...</sup>` / `<sub>...</sub>`

Token 估算由 `estimate_tokens()` 完成，用于切分和合并判定（不是模型精确 token）。

### 2.4 `parent_headings`

- 为当前块提供祖先标题链（由 `current_heading_stack` 维护）。
- 长块锚点切分后，后续子块会继承原 `parent_headings`，并在非 `Preface/Uncategorized` 时附加原块 `heading`。

### 2.5 `level`

- 1-9，对应 Word 标题层级（oxml中的`outlineLvl`0表示层级1）。
- 检测优先级：段落直接 `outlineLvl` > 样式（含 `basedOn` 继承）里的 `outlineLvl`。
- `outlineLvl=9` 视为正文。

### 2.6 `table_chunk_role`

- 合法值：`none` / `first` / `middle` / `last`。
- 合并器语义：
  - `middle`：不可与任何块合并
  - `first`：只允许向后合并
  - `last`：只允许向前合并
  - `none`：无限制

当前代码落盘现实：
- 常见值是 `none` 和 `middle`。
- `first` / `last` 在当前主流程里几乎不会最终落到输出文件（首片和末片通常以“段落形态”进入后续再打包为 `none`）。

### 2.7 `table_header`（可选）

- 仅在表格分片相关块中按需出现（直接出现在“文本块行”对象内）。
- 数据来自 Word 行属性 `w:tblHeader` 对应的表头行。
- 用于给分片表格提供列语义。

---

## 3. 智能分块主流程（`fixlevel is None`）

执行顺序（代码真实顺序）：

1. 结构扫描：按标题推进块边界
2. 表格预处理：超大表按行切片
3. 长块切分：超过 token 上限时按锚点切
4. 全局小块合并：按层级自底向上收敛

### 3.1 阶段 A：标题驱动的初始切块

- 标题判定基于 `outlineLvl`，不是样式名字符串匹配。
- 遇到“触发切块”的标题时：
  - 只有前一块存在正文内容（`has_body_content=True`）才会落盘。
  - 该标题文本会作为新块首段进入 `current_paragraphs`，不是只做元数据。
- 在这一阶段，普通正文段落与“未触发分片的表格”都会累积进同一个 `current_paragraphs`，因此默认会在同一文本块内输出（除非后续阶段再拆分）。

### 3.2 阶段 B：表格智能分片（`split_table`）

触发条件：
- 表格 JSON 估算 token `> TABLE_MAX_TOKENS (5000)`。

分片策略：
- 目标片数 = `max(ceil(total/3000), ceil(total/5000))`。
- 仅按“行边界”切分。
- 若最后一片过小（`< TABLE_MIN_LAST_CHUNK_TOKENS`，当前为 1600）且并回前片后不超 5000，则合并最后两片。

与正文拼接策略：
- 首片：并入当前段落缓存（可与表前内容同块），添加`table_chunk_role` 属性
- 中间片：立即输出独立块，文本块添加 `table_header` 和 `table_chunk_role` 属性
- 末片：写回段落缓存（可与表后文本同块），文本块添加 `table_header`属性

> 1. 为了保持保持标题的语义连贯性，表格片段的中间片和末片的标题增加一级，标题为首片的表给后添加：（表格分片<序号>）。
> 2. `table_header`属性的内容来源于word的跨页表头，如果没有设置跨页表头，这不会添加`table_header`属性。

表格分片后“正文与表格是否同块”：
- 未分片表格：与前后正文同块（如果有的话）。
- 分片表格首片：与表前正文同块（如果有的话）。
- 分片表格中间片：独立成块，不与正文同块。
- 分片表格末片：末片与表后正文同块（如果有的话）。

### 3.3 阶段 C：长块切分（`split_long_block`）

触发条件：
- 当前块 `content` 估算 token `> MAX_BLOCK_CONTENT_TOKENS (8000)`。

切分逻辑：
- 先计算目标块数：`max(ceil(total/6000), ceil(total/8000))`。
- 候选锚点：非表格段落，且 `0 < len(text) <= 100` 字符；禁止对表格进行二次切分。
- 选点方式：按每个理想位置，贪心选择“距离最近”的候选锚点。
- 被选锚点段落既作为新块 `heading`，也保留在该新块正文中（不会被丢弃）。
- 若首次切分后仍有子块 >8000，会递归继续切。
- 若没有任何可用锚点，直接报错退出（不会硬切）。

### 3.4 阶段 D：小块合并（`merge_small_blocks`）

目的：
- 减少过碎分块，降低审计时上下文割裂，能够发现更多的勾稽关系错误问题
- 在不突破 `MAX_BLOCK_CONTENT_TOKENS` 的前提下，让块大小尽量靠近 `IDEAL_BLOCK_CONTENT_TOKENS`。
- 优先保持文档层级语义：确保所有文本块都以标题开始；确保文本块的标题下不会出现属于比自己的层级更高标题的正文。
，

Phase A（同级合并）：
- 按 `level` 从深到浅处理（例如 4→3→2→1），先合并最深层标题的同级文本；再跨级把底层文本吸收到浅层级标题下。
- 仅处理“当前层级且 `<6000 tokens`”的块。
- 只和相邻同级块尝试合并，合并后不得超过 8000。

Tail Absorption（尾部吸收）：
- 当某块已达 `>=6000`，若其后连续同级尾巴总量 `<1000` 且合并后不超 8000，则一次性吸收这些尾巴。
- 遇到 `table_chunk_role="middle"` 会停止吸收。

Phase B（跨级吸收）：
- 允许高层级（数字更小）吸收相邻低层级块，方向受 `table_chunk_role` 限制。

合并后字段继承：
- `heading` / `parent_headings` / `level` 继承“被保留标题”的块；
- `uuid` 取前块，`uuid_end` 取后块；
- 新块 `table_chunk_role` 统一写为 `"none"`（未被合并的 `middle` 会保留）。

**最终分块效果（总体）：**
- 块数量通常减少，单块上下文更完整。
- 大多数块会落在“接近 6000 token、且不超过 8000 token”的区间。
- 表格中间分片（`middle`）通常保持隔离，避免被错误拼接。

---

## 4. 固定标题层级分块

`--fixlevel=N`（必须写成等号形式）是“固定标题层级切块”模式：

- 标题切块规则：
  - `N>0`：仅 `level<=N` 的标题触发切块，其他标题当正文。
  - `N=0`：所有标题都触发切块。
- 关闭内容优化流程：
  - 不做表格分片（即使 >5000）。
  - 不做最终 `merge_small_blocks()`。
  - 中途 flush 旧块时不走 `split_long_block()`。

实现细节说明：
- 当前代码在“文档末尾 final flush”仍调用一次 `split_long_block()`，因此 `fixlevel` 模式下最后一个块仍可能被长块切分。这是现实现行为。

---

## 5. 阈值总表（当前常量）

| 常量 | 值 | 用途 |
|------|----|------|
| `MAX_HEADING_LENGTH` | 200 字符 | 标题长度上限（超限直接报错） |
| `MAX_ANCHOR_CANDIDATE_LENGTH` | 100 字符 | 长块切分锚点候选上限 |
| `IDEAL_BLOCK_CONTENT_TOKENS` | 6000 | 块理想大小/合并参考 |
| `MAX_BLOCK_CONTENT_TOKENS` | 8000 | 块硬上限 |
| `SMALL_TAIL_THRESHOLD` | 1000 | 尾部吸收阈值 |
| `TABLE_IDEAL_TOKENS` | 3000 | 表格分片目标 |
| `TABLE_MAX_TOKENS` | 5000 | 表格触发分片阈值 |
| `TABLE_MIN_LAST_CHUNK_TOKENS` | 1600 | 最后一片过小时尝试并回前片 |
