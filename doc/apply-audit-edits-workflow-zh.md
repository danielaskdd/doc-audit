# 编辑工作流与回退策略

本文面向维护 `apply_audit_edits.py` 的开发者，梳理“审核结果应用”在不同 `fix_action` 下的执行路径、状态流转与回退策略。

代码入口与核心模块：

- 入口类：`skills/doc-audit/scripts/apply_audit_edits.py`
- 主流程：`skills/doc-audit/scripts/docx_edit/main_workflow_mixin.py`
- 定位阶段：`skills/doc-audit/scripts/docx_edit/item_search_mixin.py`
- 编辑实现：`skills/doc-audit/scripts/docx_edit/edit_mixin.py`
- 表格编辑：`skills/doc-audit/scripts/docx_edit/table_edit_mixin.py`
- 结果映射：`skills/doc-audit/scripts/docx_edit/item_result_mixin.py`

---

## 1. 总览：单条审核项的处理骨架

`AuditEditApplier.apply()` 对每条 `EditItem` 执行 `_process_item(item)`，整体分 4 段：

1. 预处理与锚点定位
2. 文本定位（search/match）
3. 按 `fix_action` 执行编辑
4. 状态映射为 `EditResult`（成功/告警/失败）

简化流程：

```text
JSONL -> EditItem
  -> _process_item
     -> _find_para_node_by_id(uuid)
     -> _locate_item_match(...)
     -> 按 fix_action 执行 (_apply_delete/_apply_replace/_apply_manual)
     -> _build_result_from_status(...)
```

---

## 2. 定位阶段（所有 fix_action 共用）

`_locate_item_match()` 负责在 `uuid -> uuid_end` 范围内找到 `violation_text`，并返回：

- `target_para`
- `matched_runs_info`
- `matched_start`
- `is_cross_paragraph`
- `numbering_stripped`
- `early_result`（命中某些回退场景时提前返回）

### 2.1 匹配顺序（核心）

1. 单段匹配（原文视图）
2. 自动编号剥离后重试（如 `1.`、`a)`、`表1`）
3. 跨段匹配（包含 `\n`、表格 JSON 变体）
4. 表格 JSON 匹配
5. 非 JSON 表格 raw-text 匹配
6. `boundary_crossed` 时的兜底查找（先表格单元格，再正文 segment）

### 2.2 定位失败回退（Not Found / Boundary）

关键策略：`_try_not_found_markup_retry_to_range_comment()`

- 触发条件：reason 形如 `Violation text not found(...)`
- 行为：剥离 `<sup>/<sub>` 后重试匹配，命中则直接转 `_apply_manual` 做范围注释
- `fix_action == manual`：成功时通常返回“成功无告警”
- `fix_action in (delete, replace)`：成功时返回“成功+warning”（因为降级为注释）
- 剥离 `<sup>/<sub>` 后重试匹配依然失败，则把 violation_text 中换行符前面的空白字符去掉，再进行一次匹配。原因系在审核阶段LLM会因为幻觉在段落的末尾添加额外的空格，导致匹配失败。

### 2.3 混合正文/表格片段特殊处理

当 `violation_text` 含 `<table>...</table>`：

- `delete/replace`：直接 fallback comment，返回 warning
- `manual`：提取最长文本段后继续 manual 注释流程

### 2.4 定位阶段“告警成功”与“硬失败”的分界

定位失败并不总是失败退出，当前实现里存在两类结果：

- 告警成功（`success=True, warning=True`）：
  - 常见于 `Violation text not found(C)`、mixed content、markup-strip 重试后降级注释等
- 硬失败（`success=False`）：
  - 典型是最终落入 `Violation text not found(S/M)` 且 `fix_action` 非 `manual`
  - 或锚点段落 `uuid` 根本找不到

这意味着“找不到文本”在不同边界场景下，结果语义并不完全一致。

---

## 3. fix_action 分流与编辑路径

### 3.1 `fix_action = delete`

> 审核阶段已经不会产生fix_action为delete的结果，改用replace实现delete。原因是delete没有提供前后文，容易照成误删。

入口：`_apply_delete_or_replace('delete', ...)`

#### 路径 A：单段删除

- 调用 `_apply_delete(...)`
- 产出 `w:del + commentRange + commentReference`

#### 路径 B：跨段删除（非表格）

- 条件：`is_cross_paragraph=True` 且实际命中跨多个段落，且不在表格模式
- 调用 `_apply_delete_cross_paragraph(...)`
- 逐段删除并统一 comment

#### 路径 C：跨单元格/跨行删除（表格）

- 条件：`_check_cross_cell_boundary(...)` 或 `_check_cross_row_boundary(...)` 为真
- 调用 `_apply_delete_multi_cell(...)`
- 行为：按 cell 独立删改，支持部分成功
  - 至少 1 个 cell 成功：整体记 `success`，失败 cell 写 `[CELL FAILED]` 注释
  - 全失败：返回 `fallback`，后续转普通 fallback comment

#### delete 主要回退触发

- 文本未命中：`FB_DEL_NO_HIT`
- 无可编辑 run：`FB_DEL_NO_RUN`
- 仅方程命中：`EQ_DEL_ONLY`（转 equation_fallback）
- 与既有修订重叠：`CF_OVERLAP`（conflict）
- 跨段但在表格中：`CP_TBL_SPAN` / `CP_DEL_TBL_MODE`（cross_paragraph_fallback）

---

### 3.2 `fix_action = replace`

入口：`_apply_delete_or_replace('replace', ...)`

#### 路径 A：单段替换

- 调用 `_apply_replace(...)`
- 使用 diff（支持 `<sup>/<sub>`）
- 生成 `w:del/w:ins`，并挂 comment

#### 路径 B：跨段替换（非表格）

- 条件：实际命中跨多个段，且不在 table mode
- 调用 `_apply_replace_cross_paragraph(...)`
- 进一步委托 `_apply_diff_per_paragraph(...)`

#### 路径 C：表格跨 cell/row 替换（优先拆分）

先判定 `is_cross_cell` / `is_cross_row`，然后按顺序尝试：

1. `_try_extract_single_cell_edit(...)`
2. 若失败，`_try_extract_multi_cell_edits(...)`
3. 对提取出的 cell 调 `_apply_replace_in_cell_paragraphs(...)`

结果策略：

- 单 cell 成功：正常 track changes
- 多 cell 部分成功：整体 `success`，成功部分落修订，失败 cell 写 `[CELL FAILED]`
- 多 cell 全失败：`cross_cell_fallback`（`CC_ALL_FAIL`）
- 无法拆分跨 cell 变更：`cross_cell_fallback`（`CC_XTRACT_FAIL`）
- 跨行且拆分失败：`cross_row_fallback`（`CR_XTRACT_FAIL`）

#### replace 主要回退触发

- 文本未命中：`FB_REP_NO_HIT`
- 无可编辑 run：`FB_REP_NO_RUN`
- 仅方程命中：`EQ_REP_ONLY`
- 试图修改图片/公式等特殊元素结构：`FB_REP_SPECIAL`
- 与既有修订重叠：`CF_OVERLAP`
- 表格跨段不支持直接替换：`CP_TBL_SPAN` / `CP_REP_TBL_MODE`

---

### 3.3 `fix_action = manual`

入口：`_apply_manual(...)`

manual 不做 track changes，只做评论标注（范围注释优先，必要时 reference-only）。

#### manual 正常路径

1. 找到命中 runs
2. 插入 `commentRangeStart/End + commentReference`
3. 评论内容：
   - 正常：`violation_reason + Suggestion`
   - 降级：前缀 `[FALLBACK]{reason}`

#### manual 回退点

- `FB_MAN_NO_HIT`：找不到命中
- `FB_MAN_NO_RUN`：无可挂载 run
- `FB_MAN_START_FAIL` / `FB_MAN_END_FAIL`：范围边界插入失败
- `FB_MAN_REF_FAIL`：reference-only 降级也失败

注意：manual 内部存在“段落顺序倒置”保护，可能直接降级为 reference-only 注释但仍返回 `success`。

---

### 3.4 未知 `fix_action`

- 直接 `_apply_error_comment(...)`
- 返回 `EditResult(False, ..., "Unknown action type: ...")`

---

## 4. 冲突重试策略（delete/replace 共享）

如果审核结果要求对同一个文本块的多处相同内容进行修改，后面条目的修改会命中前面已经修改过的条目，导致出现编辑冲突（当状态为 `conflict`，典型是 `CF_OVERLAP`），此时在剩余的文本块中继续重试。即在后续的文本块中进行匹配和应用规则：

1. 使用当前匹配位置之后的下一处同文本继续尝试
2. 优先扩大到整块 runs（`uuid -> uuid_end`）进行重搜
3. 直到成功或无更多匹配

若所有候选都冲突：

- 写 fallback comment：`[FALLBACK]Multiple changes overlap`
- 结果标记为 `success=True, warning=True`

---

## 5. 状态到结果的映射（`_build_result_from_status`）

| `success_status` | 处理方式 | 最终结果 |
|---|---|---|
| `success` | 正常完成 | success |
| `conflict` | 写 fallback comment | success + warning |
| `cross_paragraph_fallback` | 尝试转 manual；manual 再失败则普通 fallback comment | success + warning |
| `cross_cell_fallback` | 同上（先 manual） | success + warning |
| `cross_row_fallback` | 直接 fallback comment | success + warning |
| `equation_fallback` | 尝试转 manual | success + warning |
| `fallback` + manual | 写 error comment（含 fallback reason） | success + warning |
| `fallback` + delete/replace | 写 fallback comment | success + warning |
| 其他未知状态 | 写 error comment | fail |

---

## 6. 注释策略分层

系统里有 4 类“注释落地”策略：

1. 正常注释（配合 track changes 的 comment）
2. fallback comment：`_apply_fallback_comment()`
   - 格式：`[FALLBACK]{reason}\n{WHY}...{WHERE}...{SUGGEST}...`
3. error comment：`_apply_error_comment()`
   - 用于异常/未知状态
4. cell fallback comment：`_apply_cell_fallback_comment()`
   - 多 cell 部分失败场景，前缀 `[CELL FAILED]`

---

## 7. 状态原因（Status Reason）机制

由于一种编辑匹配或编辑策略失败后程序会改变策略或进行回退，为了避免后续的操作错误地获取到前面的错误状态设计了一个错误状态暂存和获取机制：

`_set_status_reason(status, code, detail)` + `_consume_status_reason(status)`

这是一次性“原因锁存”：

- 只在状态匹配时消费
- 消费后清空（one-shot）
- 后写覆盖前写（last-write-wins）
- 返回展示通常为：`CODE: summary`

这使日志和最终 warning 文案可携带可检索的短码（如 `CC_XTRACT_FAIL`）。

---

## 8. 失败项输出与重试

`apply()` 后，`save_failed_items()` 会把失败条目输出为 `<input>_fail.jsonl`：

- 含原 meta + 失败统计
- 每条失败项附 `_error`

适合“人工修正后重跑”：

```bash
python skills/doc-audit/scripts/apply_audit_edits.py xxx_fail.jsonl --skip-hash
```

---

## 9. 维护建议（针对 fix_action 路径）

1. 新增 `fix_action` 时，必须同时更新：
   - `_process_item` 分流
   - `_build_result_from_status` 映射
   - 对应测试（成功、fallback、异常）
2. 涉及表格替换逻辑时，优先保持“可拆分就拆分、不可拆分就降级注释”的策略，避免硬失败。
3. 涉及 `status_reason` 的新状态码，建议统一前缀：
   - `FB_*` 通用 fallback
   - `CP_*` 跨段
   - `CC_*` 跨单元格
   - `CR_*` 跨行
   - `CF_*` 冲突
