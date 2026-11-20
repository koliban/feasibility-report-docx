# editor_sdk 操作手册

基于 `edsdk.py` 的本地 Word 编辑通道。命令形态：

```bash
python edsdk.py call <tool_name> <key>=<value> ... [--json '<json>'] [--json-file <path>]
```

`edsdk.py` 位于 `workbuddy-local-office-edit` 技能目录（`docops.py doctor` 会自动探测）。所有 `idx / begin / end` 都是 **DOC 坐标（UTF-16 code unit）**，只能取自查询工具的返回值，**严禁按字符数、段落序号或行号手算**。

---

## 一、打开与保存

```bash
# 打开（wait=true 阻塞到可编辑，避免下一步读不到内容）
edsdk.py call open_file file_path="<docx路径>" wait=true
# 保存回原路径
edsdk.py call save_file file_id="<docx路径>"
```

`file_id` 可以直接用文件绝对路径，不必另存句柄。

**这两个工具返回纯文本，不是 JSON**，别拿 `ok` 字段判成败：

- `open_file` → `open completed (doc), model ready for file_id=...`
- `save_file` → `File saved to: ...`

**不要重复 `open_file`。** 对已打开的文件再调一次（默认 fresh open）会用磁盘内容重建编辑器，内存里未保存的编辑**全部丢弃**，表现为"改完没生效，重读还是老内容"。正确做法：先用 `doc_resolve_document_structure limit=1` 探测，能读到就复用。

**同一文件的正斜杠与反斜杠写法会注册成两个不同 file_id**，换写法会撞锁（`10201`）并报 `document is not open`。固定一种写法，或两种都探测一遍。

同一文件默认独占锁，被 Word/WPS 占用时报 `10201`。

---

## 二、读结构（最常用）

```bash
# 全量节点（必须 limit=0，默认 150 会静默截断尾部）
edsdk.py call doc_resolve_document_structure file_id="<docx>" mode=compact limit=0

# 只看标题骨架
edsdk.py call doc_resolve_document_structure file_id="<docx>" mode=outline limit=0

# 定位单个坐标（自动 full 回包，跳过分页）
edsdk.py call doc_resolve_document_structure file_id="<docx>" idx=7700
```

节点字段：

| 字段 | 含义 |
|---|---|
| `paragraph_index` | 全局 1-based 段落序号，**只适合人看，不要用来定位写操作** |
| `paragraph_id` | 段落 ID，用于 `doc_apply_named_style`；重开文档后失效，需重新取 |
| `start_index` / `end_index` | DOC 坐标，`start_index` 是删段/换行的入口 |
| `type` | Paragraph / Heading / Title / Subtitle / Table |
| `heading_level` | 1..9，标题层级 |
| `table_id` | 表格的稳定标识，8 字符，**跨编辑有效** |

`mode` 三档：`outline`（标题骨架）→ `compact`（10 字符 preview，够定位）→ `full`（含 run、单元格）。日常用 `compact`，要看格式才用 `full`。

表格单独查更轻：

```bash
edsdk.py call doc_list_tables file_id="<docx>"
edsdk.py call doc_get_table_info file_id="<docx>" table_id="<tid>"
```

---

## 三、查找

```bash
edsdk.py call doc_find file_id="<docx>" text="<关键词>"
```

返回 `locations[]`，每项含 `begin` / `end` / `related_text`（前后各 30 字上下文，匹配处加 `[]`）/ `paragraph_id` / 表格内匹配时的 `table_id, row, col`。

`begin` / `end` 可直接喂给 `doc_replace_text` 的 `ranges`。**不跨段落匹配**——跨段要分两次查。

注意 `doc_find` 只搜正文，**页眉页脚、文本框不在范围内**。

---

## 四、改写文字

```bash
# 按坐标替换（半开区间 [begin, end)）
edsdk.py call doc_replace_text file_id="<docx>" text="<新文本>" --json '{"ranges":[{"begin":7700,"end":7886}]}'

# 全局替换（工具名先确认：doc_find_and_replace）
edsdk.py call doc_find_and_replace file_id="<docx>" old_text="<旧>" new_text="<新>"
```

替换顺序：**长串先、短串后**。否则「XX系统建设项目」会被「XX系统」先切成碎片，剩下的「建设项目」就替换不掉了。

替换前先 `doc_find` 数一下命中数，替换后再数一次，确认清零。

---

## 五、插入内容

### 5.1 插入 Markdown 批量段落

```bash
edsdk.py call doc_insert_markdown file_id="<docx>" idx=<坐标> markdown="file://<本地 md 绝对路径>"
```

- 支持 `#` 标题、`-` 列表、`1.` 有序、`**加粗**`、表格。
- 返回 `last_edit_index` 与 `position`。**`position` 是文档末尾可操作位置，不是插入点**，别拿它定位刚插入的内容。

**idx 的语义是"插入点所在段落"**，同一个章节标题有三个可选坐标，结果完全不同（实测）：

| 传什么 | 结果 |
|---|---|
| 标题的 `end_index` | 正文并进标题段 → 标题被污染，且尾部多出一个空标题段 |
| 标题的 `start_index` | 正文插到标题前面 → 标题被推到正文之后 |
| 标题**下一段的 `start_index`** | 正确：标题独占一段，正文紧随其后，无残留 |

所以标准做法是：先拿到章节标题段，再取它在结构列表中的**下一个节点**的 `start_index` 作为插入点。`docops.py insert --after-heading` 就是这么做的。

### 5.2 插入空段落（拆分用）

```bash
edsdk.py call doc_insert_paragraph_with_text file_id="<docx>" idx=<段落end_index> text="" type=0
```

这是**唯一推荐**的空段插入方式，会得到一个独立普通空段。
`doc_insert_paragraph` 是底层拆分接口，在编号段末尾调用会得到「继续编号的空段」，不要用它造空段。

注意：用 `type=0` 拆段后，**分隔符之前的那个段落会退化为普通段落**，原本的标题层级丢失，必须紧接着重新套 `HEADING_n`。

### 5.3 插入带样式的单段标题

```bash
edsdk.py call doc_insert_paragraph_with_text file_id="<docx>" idx=<坐标> text="<标题>" type=1 level=2
```

一步完成写入 + 定级，返回 `paragraph_id`，比「先写文本再定级」少一次坐标漂移机会。

---

## 六、删除段落

```bash
edsdk.py call doc_delete_paragraph file_id="<docx>" idx=<start_index>
```

规则：

1. `idx` 传该段 `start_index`（不能传 `paragraph_index`，不能手算）；
2. **删多段必须从后往前**（`start_index` 降序），否则后面的坐标全部失效；
3. 每删一段，其后的全部坐标左移，所以下一段的原坐标**依然有效**（因为它在被删段之后、尚未变动）——这是倒序删除成立的原因；
4. 不能删文档最后一个段落。

---

## 七、样式与格式

```bash
# 按 paragraph_id 套命名样式（HEADING_1..9 / TITLE / SUBTITLE / NORMAL_TEXT）
edsdk.py call doc_apply_named_style file_id="<docx>" --json '{"paragraph_ids":["7FB57C61"],"named_style_type":"HEADING_2"}'

# 按 range 套样式（坐标已知时更方便）
edsdk.py call doc_apply_named_style file_id="<docx>" --json '{"ranges":[{"begin":1187,"end":1350}],"named_style_type":"NORMAL_TEXT"}'

# 读写文本属性（字体、字号、加粗……）——注意工具名是 update / get
edsdk.py call doc_get_text_property file_id="<docx>" idx=3300
edsdk.py call doc_update_text_property file_id="<docx>" --json '{"ranges":[{"begin":7700,"end":7886}],"text_format":{"font_family":"仿宋","font_size_pt":12}}'

# 段落属性
edsdk.py call doc_get_paragraph_property file_id="<docx>" idx=8533

# 格式刷：从一段拷格式到多段
edsdk.py call doc_copy_format file_id="<docx>" --json '{"source_range":{"begin":3120,"end":3140},"target_ranges":[{"begin":880,"end":905}]}'
```

取基准格式时**从模板原始正文段读一份**（`doc_get_text_property`），再套到新段落上。不要凭经验填「宋体小四」。

---

## 八、表格

```bash
edsdk.py call doc_list_tables file_id="<docx>"
edsdk.py call doc_get_table_info file_id="<docx>" table_id="<tid>"

# 整表填内容（一次传全部单元格，含表头）
edsdk.py call doc_set_table_cells file_id="<docx>" table_id="<tid>" --json-file cells.json
```

`cells` 结构：

```json
{
  "common_cell_properties": {"text_format": {"font_family": "仿宋", "font_size_pt": 12}},
  "cells": [
    {"row": 1, "col": 1, "text": "序号", "horizontal_align": "center"},
    {"row": 2, "col": 1, "text": "1"}
  ]
}
```

`row` / `col` 都是 **1-based**。单元格多于表格时会被忽略，少于表格时残留旧内容——**填完必须回读核对**。

行数调整：`doc_insert_table_row` / `doc_delete_table_row`（`row` 1-based）。

---

## 九、验证（不要用截图）

```bash
# 导成 HTML，检索全文（比结构树更适合查正文残留）
edsdk.py call doc_convert_to_html file_id="<docx>" output_path="<out.html>"

# 与模板对比结构与格式差异（两篇都要先 open_file）
edsdk.py call doc_compare_documents file_id="<新稿>" other_file_id="<模板>"
```

`doc_compare_documents` 返回段落级 `added / deleted / modified_content / modified_format` 差异，还能比页眉页脚、表格、批注。用它确认「章节数没变、层级没变、表格没多没少」比人肉数标题靠谱。

`doc_to_image` 依赖云转换服务，常不可用，路径含中文时更容易失败。**不要把它当验证手段**，失败就跳过，改用上面的方式。

---

## 十、常见报错

| 报错 | 处置 |
|---|---|
| `10201 file is occupied` | 先用 Word/WPS 关掉该文件，或 `open_file allow_shared=true`（无并发保护，慎用）；也可能是换斜杠写法撞了同一文件的另一把锁 |
| `document is not open` | 该 file_id 没有被打开；或换过路径写法。先探测再 `open_file` |
| 未找到工具 `doc_set_text_property` | 正确名是 `doc_update_text_property`；先 `edsdk.py list doc` 确认 |
| 结构树少了尾部节点 | 忘了 `limit=0` |
| 替换后文字被拆碎 | 长短串替换顺序反了 |
| 删除后全乱 | 未降序删 |
| 保存被拒绝、提示格式变化 | 源文件是非 OOXML 的老格式（.doc/.xls），先另存为 .docx |
| `连接失败（10061）editor_sdk 是否在运行` | 后台服务挂了。在编辑器里重新打开任意 Office 文件即可拉起；**每改完一章就 save**，避免丢工作 |

---

## 十一、服务会挂

`editor_sdk` 是本机后台服务。同时打开多份文档、或长时间高频调用，它会静默退出（报 `WinError 10061` / `WinError 10054`）。

- **每完成一章就 `save_file`**，这是唯一可靠的止损手段；
- 挂掉后重新打开任意 Office 文件即可拉起，不必重启宿主程序；
- 已在内存但未保存的改动无法找回。
