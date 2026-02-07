# docx_edit 模块划分说明

为了将 `apply_audit_edits.py` 拆分成可维护的结构，按照“职责 + 编辑场景”进行分类：

1. `common.py`：基础常量、数据结构、文本归一化工具函数。
2. `navigation_mixin.py`：Word DOM 遍历、段落/表格范围定位、run 收集与匹配。
3. `table_edit_mixin.py`：表格单元格场景的编辑提取与应用（多单元格/单单元格）。
4. `revision_mixin.py`：修订（track changes）底层算法与差异应用（diff、del/ins 构造等）。
5. `workflow_mixin.py`：删除/替换主流程（跨段落与普通段落）及变更 ID 管理。
6. `comment_workflow_mixin.py`：批注、手工处理（manual）、条目处理与整体 apply/save 流程。

这样可以快速定位函数：
- **查“怎么定位 Word 内容”**：`navigation_mixin.py`
- **查“表格编辑怎么做”**：`table_edit_mixin.py`
- **查“修订 XML 怎么生成”**：`revision_mixin.py`
- **查“处理一个 edit item 的控制流程”**：`comment_workflow_mixin.py`
