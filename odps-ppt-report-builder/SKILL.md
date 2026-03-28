---
name: odps-ppt-report-builder
description: "用固定 SQL 从 ODPS 取数、整理表结构，并按固定模板生成 PPT 的 skill。适合周报、复盘和经营分析这类重复报告。"
---

# ODPS 转 PPT

这个 skill 就做一件事：按固定 SQL 从 ODPS 取数，然后按固定版式出 PPT。

默认先照顾不懂技术的同事，所以优先走少安装、少折腾的路径。机器上如果刚好有 Python，再走脚本模式把流程跑得更稳。

## 先记住这几条

- 硬要求只有 `odpscmd`
- 不要默认用户有 Python、git、Node.js
- Qoder Work 和 Qoder IDE 分开看，不要默认一个装好另一个也能用
- 依赖不要装进 Qoder Work 程序目录
- 出 PPT 时优先原生文本框、表格、图表和形状，不要整页截图
- 除非明确需要插图，否则不要把正文重点做成位图素材
- 真正敏感的 SQL、表名、字段解释、PPT 规则，都放到 `assets/company_private_brief.md`

## 装在哪里

1. skill 本体

- 放到 `~/.qoderwork/skills/odps-ppt-report-builder/`

2. 如果需要 Python

- macOS：`~/.qoderwork-runtime/python/odps-ppt-report-builder/`
- Windows：`%USERPROFILE%\.qoderwork-runtime\python\odps-ppt-report-builder\`

3. 机器级工具

- `Python`、`odpscmd`、`Node.js`、`git`、`PowerPoint`、`WPS` 都装到用户或系统环境
- 不要装进 skill 目录，也不要装进 Qoder Work 程序目录

## 什么时候用这个 skill

- 固定 SQL 跑周报、日报、月报
- 先看表结构，再写汇报
- 只换日期或业务参数，重新出同一套报告
- 经营分析、活动复盘、管理汇报这类固定模板 PPT

## 默认流程：先走少安装模式

1. 先确认当前产品

- 如果是 Qoder Work，用 `~/.qoderwork/skills/`
- 如果是 Qoder IDE，不要默认路径一样

2. 再确认系统

- macOS：优先 `python3`
- Windows：优先 `py -3`

3. 看机器上已经有什么

- `odpscmd` 在不在
- Python 在不在
- 能不能直接生成 `.pptx`

4. 先读私有说明

- 如果 `assets/company_private_brief.md` 已经有了，就以它为准
- 如果没有，就用 `assets/company_private_brief.template.md` 做底稿，让维护人回公司机器补齐
- SQL、表结构规则、保密要求、PPT 写法都应该写在这个文件里

5. 取数时只用私有说明里写明的命令

- 不要自己猜 SQL
- 原始输出单独存到一个任务文件夹里，例如：
  - `raw_schema.txt`
  - `raw_query_overview.txt`
  - `raw_query_detail.txt`

6. 再整理成中间文件

- `report_context.md`
- `slide_plan.generated.json`

7. 出最终结果

- 能出 `.pptx` 就直接出
- 如果当前机器不方便出 `.pptx`，就给完整讲稿版本：每页标题、要点、表格内容、备注都写清楚，后面可以直接贴到 PowerPoint 或 WPS

8. 回结果时说清楚

- 文件路径
- 3 到 5 条重点结论
- 哪些数据没拿到，或者哪一步卡住了

## 如果机器上有 Python，再走脚本模式

只在两种情况下走这条路：

- 机器上已经有 Python
- 用户明确希望流程更稳、更可复用

1. 从 `assets/report_config.template.yaml` 开始
2. 初始化共享运行时

macOS:

```bash
python3 scripts/bootstrap_env.py
```

Windows:

```powershell
py -3 scripts\bootstrap_env.py
```

3. 生成报告中间包

macOS:

```bash
~/.qoderwork-runtime/python/odps-ppt-report-builder/venv/bin/python scripts/build_report_bundle.py --config <config-path>
```

Windows:

```powershell
%USERPROFILE%\.qoderwork-runtime\python\odps-ppt-report-builder\venv\Scripts\python.exe scripts\build_report_bundle.py --config <config-path>
```

4. 生成首版 slide plan

- 优先让 Qoder Work 读 `report_context.md` 后补 `slide_plan.generated.json`
- 如果模型偏弱，就先跑 `scripts/fill_slide_plan.py`，再人工补

5. 渲染结果

- `scripts/render_ppt.py`：生成 `.pptx`
- `scripts/render_preview_html.py`：生成 HTML 预览
- `scripts/render_preview_png.py`：生成 PNG 预览
- 如果 `.pptx` 渲染失败，至少把 plan 和讲稿交付出来

## 写的时候别踩这些坑

- 结论只写数据支持的内容，不要自己脑补业务原因
- 能查 schema 就先查 schema，不要猜字段含义
- 在少安装模式下，不要逼非技术同事折腾 Python
- 公共 skill 里不要硬编码敏感 SQL、表名、字段口径和写作要求
- 可复用的东西放到 `config` 和脚本里，不要散落在临时 prompt 里
- SQL、大纲、格式规则，以配置和私有说明为准

## 参考文件

- `references/workflow.md`
- `references/slide_plan_format.md`
- `references/setup_strategy_zh.md`
- `assets/company_private_brief.template.md`
- `assets/qoderwork_first_run_prompt.template.md`
