---
name: feasibility-report-docx
description: 基于既有 Word 可研报告模板，重写章节正文与表格，产出新项目的可行性研究报告 .docx。当用户给出一份可研报告模板（.docx）以及本项目建设内容（报价单、需求说明书、建设清单等）并要求「按这个模板写本项目的可研报告」时使用。也适用于基于既有长文档模板派生其他方案/报告。
agent_created: true
---

# 基于模板编制可研报告

## 用途

把一份现成的可研报告 `.docx` 当作**骨架与样式来源**，用新的项目建设内容重写全部章节正文与表格，输出一份可直接提交的新可研报告。

模板提供的是：章节层级、编号体系、页面设置、字体字号、封面与目录。模板**不提供内容**——除国家/行业标准、法规条文等固定引用外，所有业务描述必须按本项目重写。

## 何时使用

触发条件（需同时具备）：

1. 用户提供一份 `.docx` 模板，或明确说「按这份报告的格式/结构写」；
2. 用户同时提供本项目的建设内容来源（报价清单、需求说明书、建设方案、模块清单等）；
3. 产物是 Word 文档。

若用户只要提纲或纯文字稿，先出 Markdown 确认，再落 Word。

## 核心心法

1. **结构优先于文字。** 先把标题树、表格行列、段落数量改对，再灌文字。标题层级一旦错乱，目录、编号、交叉引用全部失效，返工成本远高于重写。
2. **从后往前改。** 任何编辑都会让其后的字符坐标漂移，而定位全靠 `start_index`。倒序编辑是唯一不必反复重算坐标的做法。
3. **每步改完立刻验，不要攒到最后。** 每完成一个章节，验三件事：标题有没有被污染、有没有冒出新空标题、下一个待改章节的坐标是否还是你以为的那个。
4. **表格先改。** 表格用 `table_id` 定位，不受坐标漂移影响，永远第一个动。
5. **不要靠截图验证。** 云渲染常不可用；用文档结构 + 关键词检索验证，速度快且确定。

## 工作流

### 步骤 1 · 摸清三份材料

| 材料 | 要拿到什么 | 用什么 |
|---|---|---|
| 模板 docx | 标题树（层级 / 顺序 / 每章段落数）、表格清单、封面与页眉字样 | `structure --headings`、`structure --tables` |
| 建设内容（xlsx 报价单 / 需求书） | 模块清单、模块顺序、金额或工作量权重、技术指标承诺 | `docops sheet`、`locate` |
| 参考文档（同类项目方案） | 行业政策依据、术语口径、模块业务描述的分层写法 | 只取写法与口径，**不搬内容** |

模板结构必须**完整**取回。先跑 `structure --json-out nodes.json`（内部已带 `limit=0`），确认节点总数与文档实际规模相称，再开始定位。

### 步骤 2 · 建副本，先做无坐标风险的替换

```bash
cp "<模板>.docx" "<新项目名>可研报告.docx"
```

先做全局文字替换（项目名、系统名、产品名），再改内容：

```bash
docops replace-all --file "<副本>" --old "<模板旧项目名>" --new "<本项目名>" --dry-run
```

用 `--dry-run` 先看命中数量与位置。注意替换顺序：**长串在前、短串在后**（先替换「XX系统建设项目」再替换「XX系统」），否则长串会被短串拆碎。替换完用 `locate` 复查旧名称是否清零。**封面、页眉页脚里的字样不在结构节点中**，必须单独查。

### 步骤 3 · 先改表格

`structure --tables` 拿到每个表的 `table_id`、行列数。用 `doc_set_table_cells` 整表覆盖，一次传全部单元格（含表头），并指定 `common_cell_properties.text_format`，保证字体与模板一致。

表格内容与建设内容要严格对应：行数不够就补行、多出就删行。表格残留模板业务痕迹是最常见的穿帮点。

### 步骤 4 · 逐章重写正文（从文档最后一章开始）

每一章按固定五步走：

1. `structure --range <该章起止 pi>` 拿到当前每个段落的 `pid / start_index / end_index`；
2. 把新内容写成 Markdown 文件（**一章一个文件**，便于返工）；
3. 按 `start_index` **降序**删除该章旧正文段落，保留章节标题段；
4. `insert --after-heading "<章节标题>" --md content.md` 在标题后插入新内容；
5. 重跑 `structure --range` 复查，再进入下一章。

**插入点必须用 `--after-heading`，不要手工填 idx。** `doc_insert_markdown` 的 idx 是"插入点所在段落"：

| 传什么 | 结果 |
|---|---|
| 标题的 `end_index` | 正文并进标题段，标题被污染 |
| 标题的 `start_index` | 正文塞到标题前面，标题被推到后面 |
| **标题下一段的 `start_index`** | 正确，标题与正文各自成段 |

`--after-heading` 会自动找到该标题段落，取其下一段的 `start_index`，并优先选中"类型是 Heading"的那一处匹配（同名文字在正文和目录里可能重复出现）。

### 步骤 5 · 修复标题污染（万一插入点选错）

症状：标题和正文粘成一段，标题变成「主要结论与建议一、可行性分析结论……」；或只剩一个空的标题段落悬在正文之后。

不需要手工拆，一条命令包办「删空段 → 拆段 → 复原标题文本 → 恢复 `HEADING_n`」：

```bash
# 先预演，确认判定正确
docops fix-heading --file "<副本>" --at "<标题完整文本>" --begin <start_index> --level 2 --dry-run
# 再执行
docops fix-heading --file "<副本>" --at "<标题完整文本>" --begin <start_index> --level 2
```

`--begin` 是标题文本的起始坐标（从 `structure` 或 `locate` 取），用来在同名多处命中时锁定正确的那一处，强烈建议带上。

### 步骤 6 · 每章结束就保存

```bash
docops save --file "<副本>"
```

`editor_sdk` 是本机后台服务，同时打开文档过多或长时间高频调用会挂掉（表现为"连接失败 10061"）。**每完成一章就 save**，挂掉只丢当前章未保存的部分，不必从头来。服务停止后在编辑器里重新打开任意 Office 文件即可拉起。

### 步骤 7 · 统一正文排版

模板正文段落有自己的字体字号，新插入的段落不带。两种处理：

- **保留模板默认**：什么都不做，走模板「正文」样式，通常可接受；
- **强制统一**：`structure --json-out nodes.json` 圈出新增正文段落区间，用 `doc_update_text_property` 批量设为模板既有值。**先从模板原始正文段上读一份基准值**（`doc_get_text_property`）再套用，不要凭经验填「宋体小四」。

### 步骤 8 · 收尾自查

跑 `docops verify`，逐项清零：

| 检查项 | 通过标准 |
|---|---|
| 空标题 | 0 个 |
| 标题文本长度 | 无标题超过 40 字（超长 = 被正文污染） |
| 标题层级连续性 | 每章下无跳级（H1 → H2 → H3） |
| 旧业务关键词残留 | 0 处（含封面、页眉页脚） |
| 章节数 / 表格数 | 与模板一致（除非本次业务确实需要增删） |
| 建设内容覆盖 | 每个模块都有对应小节，无遗漏、无多余 |

最后 `docops save`。

## 交付时必须告知用户的三件事

1. **目录页码不会自动更新** —— 打开 Word 后 `Ctrl+A` → `F9` 更新域。
2. **封面日期**沿用了模板，需要按实际立项月份改。
3. **金额与工期**是否已对齐本次建设内容（若报价清单有金额，必须逐项核对到章节）。

## editor_sdk 的坑（必读，能省掉大半返工）

| 现象 | 原因 | 处置 |
|---|---|---|
| 尾部节点凭空消失，按 pi 定位全错 | 结构查询默认 `limit=150` | 一律 `limit=0`；`docops` 已内置 |
| shell 里 JSON 被截断成半句，解析报错 | stdout 有长度上限 | JSON 落盘再解析，不要读 stdout |
| 插入后标题被正文污染 | idx 传了标题的 `end_index` | 用 `insert --after-heading`；已污染用 `fix-heading` |
| 正文跑到标题前面去 | idx 传了标题的 `start_index` | 同上 |
| 改完没生效，重读还是老内容 | 每次调用都重新 `open_file`，用磁盘内容重建了编辑器，内存修改被丢弃 | 先探测是否已打开，已打开就复用（`docops` 已内置） |
| 换种斜杠写法就报 `document is not open` | 正/反斜杠会注册成两个 file_id，另一个被锁 | 固定一种写法，或让 `docops` 自动探测 |
| 插入内容字体不对 | 新段落不带模板样式 | `doc_apply_named_style` / `doc_update_text_property` |
| 删除后位置全乱 | 未倒序执行 | 降序传 `start_index` |
| 用返回值 `position` 定位刚插入的内容 | `position` 是文档末尾可操作位置，不是插入点 | 从 `last_edit_index` 附近或重新 `structure` 找 |
| `open_file` / `save_file` 看不出成功没成功 | 这两个返回纯文本（`open completed` / `File saved to`），不是 JSON | 按文本关键字判成功 |
| 突然全部报"连接失败 10061" | editor_sdk 后台服务挂了 | 每章 save；重新打开任意 Office 文件拉起服务 |
| `doc_to_image` 失败或超时 | 云转换服务不可用 / 路径含中文 | 放弃渲染验证，改用结构 + 检索 |
| 标题样式丢了 | 段落拆分后样式未继承 | 按 `paragraph_id` 或 range 重新 `HEADING_n`；`doc_set_text_property` 不存在，正确名是 `doc_update_text_property` |
| 替换后文字被拆碎 | 长短串顺序颠倒 | 长串先替换 |

详细命令清单与参数见 `references/editor_sdk_playbook.md`。

## 各章写什么

模板通常是五章结构：项目概述 → 总体建设方案 → 项目实施周期 → 经济与社会效益 → 主要结论与建议。每章的必备要素、各章节的落笔顺序、常见空洞写法与替代写法见 `references/fsr_outline.md`。

行文体例（视角、称谓、禁用词、数字与单位写法）见 `references/writing_rules.md`。

## 脚本

`scripts/docops.py` —— editor_sdk 的薄封装，专治上面这些坑。

```bash
python scripts/docops.py doctor                                    # 探测 editor_sdk
python scripts/docops.py structure --file "<docx>" --headings
python scripts/docops.py structure --file "<docx>" --range 60-90
python scripts/docops.py structure --file "<docx>" --tables
python scripts/docops.py locate   --file "<docx>" --text "<关键词>"
python scripts/docops.py delete   --file "<docx>" --idx 8012,7990,7944
python scripts/docops.py insert   --file "<docx>" --after-heading "<章节标题>" --md content.md
python scripts/docops.py fix-heading --file "<docx>" --at "<标题>" --begin 7893 --level 1 [--dry-run]
python scripts/docops.py fill     --file "<docx>" --begin 7700 --end 7886 --file-in content.md
python scripts/docops.py split    --file "<docx>" --at "<标题>" --begin 7893     # fix-heading 的底层原语
python scripts/docops.py style    --file "<docx>" --text "<标题>" --style HEADING_2
python scripts/docops.py replace-all --file "<docx>" --old "<旧名>" --new "<新名>" --dry-run
python scripts/docops.py verify   --file "<docx>" --forbid "<旧业务词1>,<旧业务词2>"
python scripts/docops.py save     --file "<docx>"
python scripts/docops.py sheet    --file "<报价单.xlsx>" --rows 0-40
```
