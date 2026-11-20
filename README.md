# 可研报告 Word 编制技能

根据既有的 Word 可研报告模板和新项目的建设材料，重写章节正文、更新表格，并生成新的 `.docx`。输入材料可以是需求说明书、报价单或建设清单。技能提供章节提纲、写作规则和文档检查命令。

## 安装到 Codex

让 Codex 从 GitHub 仓库 `koliban/feasibility-report-docx` 安装 `skills/feasibility-report-docx`，下一轮调用 `$feasibility-report-docx`。提供 Word 模板和本项目的建设资料后，可以这样发起任务：

> 用 $feasibility-report-docx 根据这份 Word 模板和建设清单编制本项目的可行性研究报告，保留模板的章节层级和排版。

## 运行环境

命令行脚本依赖 Python 3 和 WorkBuddy 的本地 Office 编辑服务。`docops.py` 会查找常见 Windows 安装位置中的 `edsdk.py`，也支持将环境变量 `EDSDK_PATH` 指向已有的 `edsdk.py`。启动 WorkBuddy 并打开一份 Office 文档后，可以从仓库根目录检查环境和文档结构：

```powershell
python skills/feasibility-report-docx/scripts/docops.py doctor
python skills/feasibility-report-docx/scripts/docops.py structure --file "C:/path/to/template.docx" --headings
```

使用自己的模板和建设材料。仓库包含技能说明、参考资料和 Python 脚本，WorkBuddy 和 `editor_sdk` 需另行获取。

- [技能工作流](skills/feasibility-report-docx/SKILL.md)
- [可研报告章节提纲](skills/feasibility-report-docx/references/fsr_outline.md)
- [写作规则](skills/feasibility-report-docx/references/writing_rules.md)
- [本地编辑服务操作手册](skills/feasibility-report-docx/references/editor_sdk_playbook.md)
- [MIT 许可证](LICENSE)
