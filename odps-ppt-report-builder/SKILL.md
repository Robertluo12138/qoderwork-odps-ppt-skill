---
name: odps-ppt-report-builder
description: "用固定 SQL 从 ODPS 取数、整理表结构，并按淘宝闪购品牌 VI 标准生成月度经营诊断 PPT。适合周报、月报和经营分析类重复报告。"
---

# ODPS 转 PPT

这个 skill 做一件事：按固定 SQL 从 ODPS 取数，然后按品牌 VI 标准生成经营诊断 PPT。

默认先照顾不懂技术的同事，优先走少安装路径。如果机器上有 Python，再走脚本模式。

---

## 核心原则（必须遵守）

1. 硬要求只有 `odpscmd`，不要默认用户有 Python、git、Node.js
2. Qoder Work 和 Qoder IDE 是两个独立产品，路径不同，不要混用
3. 依赖不要装进 Qoder Work 程序目录，也不要装进 skill 目录
4. 出 PPT 时用原生文本框、表格、图表和形状，**不要整页截图**
5. 敏感的 SQL、表名、字段解释、PPT 规则，都放 `assets/company_private_brief.md`
6. 一次只处理一个操作系统，**不要把 Mac 和 Windows 的命令混在一起说**
7. 每一步先回报结果，再进入下一步。不要一口气说 5 步让用户自己跑
8. 如果缺文件、缺命令、缺权限，**不要假装能继续**，直接说清楚卡在哪里

---

## 品牌 VI 标准

生成的 PPT 必须遵循以下品牌规范：

| 元素 | 标准 |
|---|---|
| 品牌色 | `#FF6200`（PANTONE Orange 016C），占比 90% |
| 辅助色 | `#FFA020`，占比 10% |
| 淡底填充 | `#FFF4EB` |
| 暖边框 | `#FFDCC8` |
| 图表配色 | 5 色梯度：`FF6200` → `FF8A3D` → `FFA020` → `FFB85C` → `FFDCC8` |
| Logo | `assets/taobao_flash_sale_logo.png`，封面左上角 + 每页页脚右下角 |
| 结构分析页 | 橙色全宽横幅在页面顶部，白色粗体字显示核心结论 |
| 图表类型 | 根据数据特征自动选择：柱状图、横条图、甜甜圈图等，不要全部用同一种 |

完整的品牌 VI 文档：`references/brand_vi_spec.md`
模板视觉说明：`references/ppt_template_catalog.md`

### 跨平台字体规范

| 平台 | 西文字体 | 中文字体 | 备用字体 |
|---|---|---|---|
| macOS | Aptos | Hiragino Sans GB | Arial Unicode MS |
| Windows | Aptos | Microsoft YaHei | Arial Unicode MS |
| Linux | DejaVu Sans | Noto Sans CJK SC | Arial Unicode MS |

上面的字体映射已经硬编码在 `report_config.template.yaml` 的 theme 部分。渲染脚本会自动检测当前操作系统并选对应字体，**不需要手动修改**。

---

## 短需求处理

用户如果只发一句话，比如：

- "帮我生成霸王茶姬三月的 PPT"
- "帮我出霸王茶姬月报"
- "帮我重跑上周周报"

按下面规则理解：

1. 从 `assets/company_private_brief.md` 找业务别名、报告模板和默认口径
2. 没写年份就默认当前年份，但在回复里**明确复述这个假设**
3. 没写报告类型就按私有说明里的默认模板来
4. 只缺一个关键信息，就只问一个短问题，不要连续追问
5. 必要信息够了就直接开始，不要先讲方案

---

## 首次检查（按顺序执行，每步都要回报结果）

### 第 1 步：检查 skill 位置

**macOS：**
```bash
ls ~/.qoderwork/skills/odps-ppt-report-builder/SKILL.md
```

**Windows（PowerShell）：**
```powershell
Test-Path "$env:USERPROFILE\.qoderwork\skills\odps-ppt-report-builder\SKILL.md"
```

如果文件不存在，告诉用户把 `odps-ppt-report-builder` 文件夹放到对应目录。

### 第 2 步：检查私有说明

```
assets/company_private_brief.md
```

如果不存在，告诉用户：

> 现在缺的是：`assets/company_private_brief.md`  
> 你下一步只需要做这件事：复制 `assets/company_private_brief.template.md`，改名为 `company_private_brief.md`，让维护人补真实 SQL 和业务规则。  
> 做完以后发"继续"给我。

**不要给开放式问题，不要说"你想怎么做"，不要把 3 条路一起丢给用户选。**

### 第 3 步：检查 odpscmd

**macOS：**
```bash
which odpscmd && odpscmd --version
```

**Windows（PowerShell）：**
```powershell
where.exe odpscmd; odpscmd --version
```

如果找不到，告诉用户联系数据组安装 odpscmd 并配置好。

### 第 4 步：检查 Python（可选）

**macOS：**
```bash
python3 --version
```

**Windows（PowerShell）：**
```powershell
py -3 --version
```
如果上面不行，再试：
```powershell
python --version
```

- **没有 Python**：继续走少安装模式，不要逼非技术同事装 Python
- **有 Python**：走脚本模式，效果更好

### 第 5 步：检查 Logo 文件

```
assets/taobao_flash_sale_logo.png
```

如果存在就会自动带上品牌 Logo。如果不存在，PPT 会正常生成但没有 Logo。

---

## 少安装模式（默认路径，不需要 Python）

适合非技术同事。步骤：

1. 确认 `odpscmd` 可用
2. 读取 `assets/company_private_brief.md` 获取 SQL 和口径
3. 用 `odpscmd` 执行 SQL，把原始结果保存到任务文件夹
4. AI 分析数据后生成 `report_context.md` 和 `slide_plan.generated.json`
5. 如果机器没法渲染 `.pptx`，就把完整讲稿交付（每页标题、要点、表格、备注全部写清楚）
6. 回结果时说清楚：文件路径 + 3~5 条重点结论 + 哪些数据没拿到

---

## 脚本模式（有 Python 时使用）

### 第 1 步：初始化环境

**macOS：**
```bash
cd ~/.qoderwork/skills/odps-ppt-report-builder
python3 scripts/bootstrap_env.py
```

**Windows（PowerShell）：**
```powershell
cd "$env:USERPROFILE\.qoderwork\skills\odps-ppt-report-builder"
py -3 scripts\bootstrap_env.py
```

> 这一步会在 `~/.qoderwork-runtime/python/odps-ppt-report-builder/venv/` 创建虚拟环境并安装依赖。  
> 如果 `py -3` 不可用，改用 `python scripts\bootstrap_env.py`。

### 第 2 步：准备配置

1. 复制 `assets/report_config.template.yaml` 到任务目录
2. 修改里面的 `vars` 部分：`brand_name`、`start_date`、`end_date` 等
3. 确认 `theme` 部分的颜色和字体是否需要调整（一般不需要）

### 第 3 步：取数

**macOS：**
```bash
~/.qoderwork-runtime/python/odps-ppt-report-builder/venv/bin/python \
  scripts/build_report_bundle.py --config <config-path>
```

**Windows（PowerShell）：**
```powershell
& "$env:USERPROFILE\.qoderwork-runtime\python\odps-ppt-report-builder\venv\Scripts\python.exe" `
  scripts\build_report_bundle.py --config <config-path>
```

产出：`report_payload.json`（查询结果）+ `report_context.md`（数据摘要）

### 第 4 步：生成 slide plan

优先让 AI 读 `report_context.md` 和 `assets/company_private_brief.md` 后生成 `slide_plan.generated.json`。

如果模型能力弱，也可以让 AI 先跑：
```bash
# macOS
~/.qoderwork-runtime/python/odps-ppt-report-builder/venv/bin/python \
  scripts/fill_slide_plan.py --config <config-path>
```

slide plan 格式定义：`references/slide_plan_format.md`

### 第 5 步：渲染 PPT

**macOS：**
```bash
~/.qoderwork-runtime/python/odps-ppt-report-builder/venv/bin/python \
  scripts/render_ppt.py \
  --config <config-path> \
  --payload <payload-path> \
  --plan <plan-path>
```

**Windows（PowerShell）：**
```powershell
& "$env:USERPROFILE\.qoderwork-runtime\python\odps-ppt-report-builder\venv\Scripts\python.exe" `
  scripts\render_ppt.py `
  --config <config-path> `
  --payload <payload-path> `
  --plan <plan-path>
```

还可以同时生成预览：

- `scripts/render_preview_png.py`：PNG contact sheet 预览
- `scripts/render_preview_html.py`：可交互 HTML 预览

如果 `.pptx` 渲染失败，至少把 plan 和讲稿交付出来。

---

## 图表类型自动选择

渲染器会根据 slide 模板类型自动选择最佳图表样式：

| 模板 | 图表类型 | 原因 |
|---|---|---|
| `weekly_trend` | 柱状图 (column) | 时间趋势对比 |
| `comparison` | 横条图 (bar) | 月度对比，类别少 |
| `time_slot` | 柱状图 (column) | 时段间并列对比 |
| `price_band` | 甜甜圈 (doughnut) | 价格带占比分布 |
| `city_distribution` | 横条图 (bar) | 城市多适合横排 |
| `aoi_distribution` | 甜甜圈 (doughnut) | AOI 占比分布 |
| `new_old_mix` | 甜甜圈 (doughnut) | 客群组成比例 |
| `cluster_mix` | 柱状图 (column) | 人群分段对比 |

如果在 slide plan 里显式指定了 `chart_type`，优先用指定的类型。

---

## 常见问题排查

### `odpscmd` 找不到

macOS：检查 `~/.bash_profile` 或 `~/.zshrc` 里有没有把 odpscmd 的路径加到 PATH。  
Windows：检查系统环境变量 PATH 里有没有 odpscmd 的安装路径。

### 字体渲染不对

确认 `report_config.template.yaml` 里 `cjk_font_family_mac` 和 `cjk_font_family_windows` 的字体是否在当前系统中安装。默认配置使用的字体（Hiragino Sans GB / Microsoft YaHei）分别是 macOS 和 Windows 自带的，正常情况下不需要额外安装。

### Logo 没有显示

检查 `assets/taobao_flash_sale_logo.png` 文件是否存在。如果不存在，渲染器会静默跳过 Logo（不会报错），PPT 其他内容正常生成。

### PPT 打开后图表是空的

确认 `report_payload.json` 里对应的 query 有数据。如果 SQL 跑出来是空的，渲染器会显示"这一页没有足够数据"的提示文字。

### Windows 上 `python` 命令不可用

依次尝试：`py -3`、`python3`、`python`。如果都不行，需要先安装 Python 3.10+。  
安装时勾选 **"Add Python to PATH"**。

---

## 写报告时别踩这些坑

- 结论只写数据支持的内容，不要自己脑补业务原因
- 能查 schema 就先查 schema，不要猜字段含义
- 在少安装模式下，不要逼非技术同事折腾 Python
- 用户只说一句话，不要回一大段解释，先补齐参数后直接执行
- `company_private_brief.md` 里的 SQL 和表名是保密的，不要在回复中暴露
- 可复用的东西放到 config 和脚本里，不要散落在临时 prompt 里

---

## 参考文件

| 文件 | 用途 |
|---|---|
| `references/workflow.md` | 完整工作流程 |
| `references/slide_plan_format.md` | slide plan JSON 格式定义 |
| `references/brand_vi_spec.md` | 品牌 VI 标准色和 Logo 使用规范 |
| `references/ppt_template_catalog.md` | PPT 模板类型和视觉方向 |
| `references/setup_strategy_zh.md` | 安装策略说明 |
| `assets/company_private_brief.template.md` | 私有业务说明模板 |
| `assets/quick_start_for_colleagues.md` | 首次使用 prompt |
| `assets/qoderwork_first_run_prompt.template.md` | 首次 prompt 模板 |
