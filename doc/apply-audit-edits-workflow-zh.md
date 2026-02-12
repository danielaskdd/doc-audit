# Word无头编辑脚本说明

本文面向维护 `apply_audit_edits.py` 的开发者，梳理“审核结果应用”在不同 `fix_action` 下的执行路径、状态流转与回退策略。

代码入口与核心模块：

- 入口类：`skills/doc-audit/scripts/apply_audit_edits.py`
- 主流程：`skills/doc-audit/scripts/docx_edit/main_workflow_mixin.py`
- 定位阶段：`skills/doc-audit/scripts/docx_edit/item_search_mixin.py`
- 编辑实现：`skills/doc-audit/scripts/docx_edit/edit_mixin.py`
- 表格编辑：`skills/doc-audit/scripts/docx_edit/table_edit_mixin.py`
- 结果映射：`skills/doc-audit/scripts/docx_edit/item_result_mixin.py`

编辑能力和限制：

* 审核结果输出的编辑条目含有：`fix_action`、`violation_text` 和 `revised_text`
* fix_action 为 manual 表示仅添加范围批注，批注包裹范围为`violation_text`
* fix_action 为 replace 表示修订，通过对比`violation_text` 和`revised_text`的差异确定删除或插入的内容

* 支持的编辑：正文、单元格、列表、标题、上标、下标

* 自动编号：支持匹配时剥离编号前缀，但不直接编辑编号实体
* 公式、图片：可定位并保留占位，不支持内容级修改/删除/新增

需要解决的麻烦事：

* 无法正确自动编号：每个文本块独立编辑，无法正确计算自动编号
* 在一条规则中进行多个表格单元格的修订
* LLM审核输出的 `violation_text` 偶尔会遇原文有出入

---

## 1. 总览：单条编辑项的处理骨架

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

### 1.1 关键对象与返回值语义（速查）

`EditItem`（输入条目）关键字段：

- `uuid`：起始段落 `w14:paraId`，用于锚点定位
- `uuid_end`：结束段落 `w14:paraId`，用于限定搜索范围（必填）
- `violation_text`：要命中的原文片段
- `revised_text`：建议文本（replace/manual 使用）
- `fix_action`：`delete | replace | manual`
- `violation_reason`：写入 comment 的主要说明文本

`EditResult`（输出结果）字段语义：

- `success`：是否处理成功（包含“成功但降级为 comment”）
- `warning`：是否发生预期降级（fallback/comment-only）
- `error_message`：失败原因或降级原因（通常包含状态码前缀）
- `item`：原始 `EditItem`

`runs_info`（内部运行态）常见字段：

- `text/start/end`：用于匹配与切分的文本及偏移
- `elem/rPr/para_elem`：DOM 定位信息
- `host_para_elem`：vMerge 等场景下的真实宿主段落
- `cell_elem/row_elem`：表格定位信息
- `is_para_boundary/is_json_boundary/is_cell_boundary/is_row_boundary`：合成边界标记（非可编辑 run）
- `original_text`：表格 JSON 模式下未转义原文（编辑时使用）

## 2. 定位阶段（所有 fix_action 共用）

`_locate_item_match()` 负责在 `uuid -> uuid_end` 范围内找到 `violation_text`，并返回：

`target_para`

- 类型：段落 XML 节点（`w:p`）或 `None`
- 含义：后续编辑或 comment 的目标段落锚点
- 典型来源：单段命中时为命中段；跨段命中时通常为 `anchor_para`；表格 JSON 命中时为表格首段
- 为空场景：未命中且提前返回（`early_result != None`）

`matched_runs_info`

- 类型：`List[Dict]` 或 `None`
- 含义：用于本次命中/编辑的 run 序列（包含 `text/start/end` 与定位元数据）
- 关键特性：可能包含合成边界 run（如 `is_para_boundary/is_json_boundary`），调用方需再过滤可编辑 run
- 为空场景：未命中且提前返回

`matched_start`

- 类型：`int`
- 含义：命中起点偏移，基于 `matched_runs_info` 拼接文本坐标系
- 典型值：`>= 0` 表示命中；`-1` 表示未命中

`violation_text`（返回态，可能被归一化/重映射）

- 类型：`str`
- 含义：定位阶段“实际用于后续编辑”的文本，不一定等于输入值
- 可能变化：编号剥离、`\n`/`\\n` 归一化、表格 JSON 归一化、命中后按原文切片重映射、mixed-content 场景取最长段

`revised_text`（返回态，replace 且编号剥离时可能同步调整）

- 类型：`str`
- 含义：后续 replace/manual 使用的建议文本
- 可能变化：当 `fix_action == replace` 且命中使用了编号剥离策略时，会同步剥离建议文本中的对应编号

`is_cross_paragraph`

- 类型：`bool`
- 含义：本次命中是否按跨段模式处理
- 注意：表格 JSON 模式通常为 `True`；raw-table-cell 命中可能为 `False`（在单单元格段内定位）
- 补充：该值表示“定位形态”，最终是否走跨段编辑分支还会再按实际命中 run 分布判断

`numbering_stripped`

- 类型：`bool`
- 含义：是否通过“去编号”变体才命中（包括正文编号和表格行号场景）
- 用途：主要用于日志/结果解释，并指导 replace 分支是否调整 `revised_text`

`early_result`（命中某些回退场景时提前返回）

- 类型：`EditResult | None`
- 含义：定位阶段已决定最终结果（例如直接 fallback/comment），主流程应立即返回
- 约束：当 `early_result is not None` 时，`target_para/matched_runs_info/matched_start` 可能无效；只有 `early_result` 应被消费
- 常见值：`EditResult(success=True, warning=True, ...)`（降级注释）；`EditResult(success=False, ...)`（定位硬失败）

返回值协同约束：

- 正常路径：`early_result is None`，且 `target_para != None`、`matched_runs_info != None`、`matched_start >= 0`
- 提前返回路径：优先使用 `early_result`，其余字段仅作调试信息，不应继续用于编辑

### 2.1 匹配顺序（核心）

1. 单段匹配（原文视图）
2. 自动编号剥离后重试（如 `1.`、`a)`、`表1`）
3. 跨段匹配（包含 `\n`、表格 JSON 变体）
4. 表格 JSON 匹配
5. 非 JSON 表格 raw-text 匹配
6. `boundary_crossed` 时的兜底查找（先表格单元格，再正文 segment）

### 2.2 定位失败回退（Not Found / Boundary）

关键策略：`_try_not_found_markup_retry_to_range_comment()`

- 触发条件：reason 形如 `violation text not found(...)`
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
  - 常见于 `violation text not found(C)`、mixed content、markup-strip 重试后降级注释等
- 硬失败（`success=False`）：
  - 典型是最终落入 `violation text not found(S/M)` 且 `fix_action` 非 `manual`
  - 或锚点段落 `uuid` 根本找不到

这意味着“找不到文本”在不同边界场景下，结果语义并不完全一致。

---

## 3. 编辑阶段（按fix_action 分流与编辑路径）

### 3.1 `fix_action = delete`

> 审核阶段已经不会产生fix_action为delete的结果，改用replace实现delete。原因是delete没有提供前后文，容易造成误删。

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
- 仅方程命中：`EQ_MODIFY`
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

## 4. 编辑冲突重试策略（delete/replace 共享）

如果审核结果要求对同一个文本块的多处相同内容进行修改，后面条目的修改会命中前面已经修改过的条目，导致出现编辑冲突（当状态为 `conflict`，典型是 `CF_OVERLAP`），此时在剩余的文本块中继续重试。即在后续的文本块中进行匹配和应用规则：

1. 使用当前匹配位置之后的下一处同文本继续尝试
2. 优先扩大到整块 runs（`uuid -> uuid_end`）进行重搜
3. 直到成功或无更多匹配

若所有候选都冲突：

- 写 fallback comment：`[FALLBACK]Multiple changes overlap`
- 结果标记为 `success=True, warning=True`

---

## 5. 编辑状态到结果的映射（`_build_result_from_status`）

`_build_result_from_status` 是“状态收口层”：把编辑阶段返回的内部状态字符串统一转换成最终 `EditResult`，并集中处理降级 comment 的副作用。

输入关注点：

- `success_status`：编辑阶段返回状态（如 `success/conflict/fallback/...`）
- `status_reason`：从状态锁存消费的原因文本（通常是 `CODE: summary`）
- `target_para/anchor_para`：comment 挂载位置
- `matched_runs_info/matched_start`：需要二次转 manual 时复用

映射矩阵（按实现）：

| `success_status` | 内部动作 | `error_message` 来源 | 最终 `EditResult` |
|---|---|---|---|
| `success` | 直接返回 | `None` | `success=True, warning=False` |
| `conflict` | 写 fallback comment（reason 固定为 `Multiple changes overlap`） | 固定文案，不使用 `status_reason` | `success=True, warning=True` |
| `cross_paragraph_fallback` | 调 `_build_manual_fallback_result`：先尝试 `_apply_manual`，失败再写 fallback comment | `status_reason` 或默认文案；manual 失败时追加 `(comment also failed)` | `success=True, warning=True` |
| `cross_cell_fallback` | 同上（先 manual，再 fallback comment） | 同上 | `success=True, warning=True` |
| `equation_fallback` | 同上（先 manual，再 fallback comment） | 同上 | `success=True, warning=True` |
| `cross_row_fallback` | 不尝试 manual，直接写 fallback comment | `status_reason` 或默认文案 | `success=True, warning=True` |
| `fallback` 且 `fix_action=manual` | 写 error comment（带 fallback reason） | `status_reason` 或 `No editable runs found` | `success=True, warning=True` |
| `fallback` 且 `fix_action in (delete, replace)` | 写 fallback comment | `status_reason` 或 `No editable runs found` | `success=True, warning=True` |
| 其他未知状态 | 写 error comment（挂到 `anchor_para`） | 固定为 `Operation failed` | `success=False, warning=False` |

设计意图：

- 把“可降级场景”统一归为 warning-success，尽量不丢任务
- 真正 fail 仅保留给未知状态或前置异常路径
- comment 副作用集中在这一层，避免分散到各编辑函数中

---

## 6. 批注（注释）策略分层

系统里有 4 类“批注（注释）落地”策略：

1. 正常comment（Manual和Repalce完成后的 comment）
2. 编辑失败后的fallback comment：`_apply_fallback_comment()`
   - 格式：`[FALLBACK]{reason}\n{WHY}...{WHERE}...{SUGGEST}...`
3. error comment：`_apply_error_comment()`
   - 用于异常/未知状态
4. cell fallback comment：`_apply_cell_fallback_comment()`
   - 多 cell 部分失败场景，前缀 `[CELL FAILED]`
   - 对部分失败的cell单独添加批注

---

## 7. 状态原因（Status Reason）机制

该机制用于在编辑流程中暂存并传递失败/降级原因，供结果映射阶段统一消费：

`_set_status_reason(status, code, detail)` + `_consume_status_reason(status, include_detail=False)`

行为规则：

- `_set_status_reason(...)` 会写入锁存对象：`{status, code, summary, detail}`
- `_set_status_reason(...)` 的返回值是 `status` 本身，便于直接 `return self._set_status_reason(...)`
- `code` 为空时会标准化为 `FB_UNKNOWN`
- `detail` 会做空白归一化；`summary` 为精简版（首句 + 截断）
- `_consume_status_reason(status)` 仅在状态匹配时返回原因文本，否则返回 `None`
- 状态匹配消费后立即清空；状态不匹配时不清空（等待后续匹配或被覆盖）
- 同一条 item 内多次写入时，后写覆盖前写（last-write-wins）
- 默认返回格式为 `CODE` 或 `CODE: summary`（`include_detail=True` 时改为 `CODE: detail`）

生命周期（按主流程）：

- 每条 item 开始时 `_reset_status_reason()` 清空旧值
- 编辑阶段各分支通过 `_set_status_reason(...)` 挂载原因
- `_process_item()` 在进入 `_build_result_from_status(...)` 前调用 `_consume_status_reason(success_status)` 一次并透传
- 少数嵌套分支（如单 cell replace 失败）会先行消费并封装成新的上层状态原因

价值：

- 让最终 warning/error 文案携带稳定短码（如 `CC_XTRACT_FAIL`）
- 避免不同 fallback 分支串用旧原因，降低误诊断概率

---

## 8. 关键函数参数与返回值（代码契约）

### 8.1 `AuditEditApplier.apply() -> List[EditResult]`

参数：

- 无显式参数（依赖实例初始化参数：`jsonl_path/output_path/author/initials/skip_hash/verbose`）

返回：

- `List[EditResult]`：与输入条目一一对应

语义要点：

- 当 `skip_hash=False` 时先校验 `source_hash`
- 内部统一初始化 comment/change id、时间戳、段落顺序缓存
- warning 条目仍为 `success=True`

### 8.2 `_process_item(item: EditItem) -> EditResult`

参数：

- `item`：单条编辑任务，核心使用 `uuid/uuid_end/fix_action/violation_text/revised_text`

返回：

- `EditResult`：包含最终成功/告警/失败结果

语义要点：

- 先按 `uuid` 找 anchor 段，失败则直接 `success=False`
- 调 `_locate_item_match(...)`；若返回 `early_result` 则提前结束
- `delete/replace` 共享 `_apply_delete_or_replace(...)`，冲突会在剩余内容继续重试
- 最终统一交给 `_build_result_from_status(...)`

### 8.3 `_locate_item_match(...) -> Dict[str, Any]`

参数：

- `item`：读取 `fix_action/uuid_end`
- `anchor_para`：由 `uuid` 找到的起始段
- `violation_text/revised_text`：当前条目的查找/替换文本

返回字段：

- `target_para`：后续编辑挂载段落（可能是 anchor，也可能是命中段/表格首段）
- `matched_runs_info`：参与命中的 run 序列
- `matched_start`：相对于 `matched_runs_info` 拼接文本的起始偏移
- `violation_text/revised_text`：可能被重写后的文本
- `numbering_stripped`：是否经过编号剥离路径
- `is_cross_paragraph`：命中是否跨段（表格 JSON 模式通常为 true）
- `early_result`：`EditResult | None`，非空时主流程直接返回

### 8.4 `_apply_delete/_apply_replace/_apply_manual` 状态返回

共同点：

- 返回 `str` 状态，不直接返回 `EditResult`
- 常见值：`success | fallback | conflict | cross_paragraph_fallback | cross_cell_fallback | cross_row_fallback | equation_fallback`
- 详细原因通过 `_set_status_reason(...)` 锁存，再由 `_consume_status_reason(...)` 消费

各函数核心参数：

- `_apply_delete(para_elem, violation_text, violation_reason, orig_runs_info, orig_match_start, author)`
- `_apply_replace(para_elem, violation_text, revised_text, violation_reason, orig_runs_info, orig_match_start, author, skip_comment=False)`
- `_apply_manual(para_elem, violation_text, violation_reason, revised_text, orig_runs_info, orig_match_start, author, is_cross_paragraph=False, fallback_reason=None)`

### 8.5 `_build_result_from_status(...) -> EditResult`

关键参数含义：

- `success_status`：编辑阶段返回的状态字符串
- `status_reason`：可选 `CODE: summary`，来自一次性状态锁存
- `target_para/anchor_para`：用于失败或降级 comment 的挂载点
- `matched_runs_info/matched_start/is_cross_paragraph`：manual 降级时复用定位结果
- `numbering_stripped`：仅用于日志提示，不改变结果语义

返回：

- 统一后的 `EditResult`（成功、告警成功或失败）

### 8.6 其他关键函数契约

`_try_not_found_markup_retry_to_range_comment(...) -> bool`

- 关键参数：`item/anchor_para/violation_text/revised_text/reason/use_fallback_reason`
- 返回 `True`：已通过“去 `<sup>/<sub>` + 可选换行前空白归一化”命中，并成功降级为范围 comment
- 返回 `False`：未命中或降级失败，调用方继续原 fallback 路径

`_collect_runs_info_across_paragraphs(start_para, uuid_end) -> (runs_info, combined_text, is_cross_paragraph, boundary_error)`

- `runs_info`：跨段（含表格 JSON 边界）run 序列
- `combined_text`：拼接后的可搜索文本
- `is_cross_paragraph`：是否跨段
- `boundary_error`：`None` 或 `boundary_crossed`

---

## 9. 失败编辑条目输出与重试

`apply()` 后，`save_failed_items()` 会把失败条目输出为 `<input>_fail.jsonl`：

- 含原 meta + 失败统计
- 每条失败项附 `_error`

适合“人工修正后重跑”：

```bash
python skills/doc-audit/scripts/apply_audit_edits.py xxx_fail.jsonl --skip-hash
```

> 重跑前应当先接受把之前的修订并清空所有批注，避免新旧批注信息叠加在一起。

---

## 10. 维护建议（针对 fix_action 路径）

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
   - `NF_*`匹配失败
