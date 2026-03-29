---
name: odps-ppt-report-builder
description: "固定 SQL 从 ODPS 取数，按品牌 VI 标准出经营诊断 PPT。"
---

# ODPS → PPT

按固定 SQL 从 ODPS 取数，按品牌 VI 出经营诊断 PPT。  
没 Python 走少安装模式，有 Python 走脚本模式。

## 规矩

- 硬要求只有 `odpscmd`，别默认用户有 Python / git / Node
- Qoder Work ≠ Qoder IDE，路径不同不要混
- 依赖不要装进 Qoder Work 程序目录或 skill 目录
- PPT 用原生文本框 / 表格 / 图表，别截图贴
- 敏感 SQL、表名、字段口径放 `assets/company_private_brief.md`
- Mac / Windows 命令分开说，别混一起
- 一步做完回报结果再做下一步，别一口气甩 5 步
- 缺东西就说缺什么，别假装能继续

## 品牌 VI

| 项 | 值 |
|---|---|
| 品牌色 | `#FF6200`，占 90% |
| 辅助色 | `#FFA020`，占 10% |
| 淡底 | `#FFF4EB` |
| 暖边框 | `#FFDCC8` |
| 图表色 | `FF6200 → FF8A3D → FFA020 → FFB85C → FFDCC8` |
| Logo | `assets/taobao_flash_sale_logo.png`，封面左上 + 每页右下 |
| 横幅 | 结构分析页顶部橙色满宽条，白字写核心结论 |

字体自动按平台选：Mac 用 Hiragino Sans GB，Windows 用 Microsoft YaHei。  
详见 `references/brand_vi_spec.md`。

## 短需求

用户说"帮我生成霸王茶姬三月 PPT"这种话时：

1. 先从 `assets/company_private_brief.md` 找别名和默认口径
2. 没写年份默认今年，回复里说清楚
3. 只缺一个信息就问一个问题，别连续追问
4. 够了就直接跑，别先讲方案

## 首次检查

按顺序做，每步回报再往下走。

**1. skill 位置**

```bash
# Mac
ls ~/.qoderwork/skills/odps-ppt-report-builder/SKILL.md

# Windows PowerShell
Test-Path "$env:USERPROFILE\.qoderwork\skills\odps-ppt-report-builder\SKILL.md"
```

**2. 私有说明**

看 `assets/company_private_brief.md` 在不在。不在就告诉用户：

> 缺 `assets/company_private_brief.md`。  
> 复制 `company_private_brief.template.md` 改名，让维护人补上 SQL 和口径。  
> 补完发"继续"。

别给开放式问题。

**3. odpscmd**

```bash
# Mac
which odpscmd

# Windows
where.exe odpscmd
```

找不到就让用户找数据组装。

**4. Python（可选）**

```bash
# Mac
python3 --version

# Windows（按顺序试）
py -3 --version
python --version
```

没有就走少安装模式，别逼人装。

## 少安装模式

不需要 Python，步骤：

1. 确认 `odpscmd` 能用
2. 读 `assets/company_private_brief.md` 拿 SQL
3. `odpscmd` 跑 SQL，原始结果存到任务文件夹
4. AI 分析数据，生成 `report_context.md` + `slide_plan.generated.json`
5. 出不了 `.pptx` 就给完整讲稿（每页标题 / 要点 / 表格 / 备注）
6. 回结果时附：文件路径 + 3~5 条结论 + 缺了哪些数据

## 脚本模式

机器上有 Python 时用这条路，出图更稳。

**初始化环境：**

```bash
# Mac
cd ~/.qoderwork/skills/odps-ppt-report-builder
python3 scripts/bootstrap_env.py

# Windows
cd "$env:USERPROFILE\.qoderwork\skills\odps-ppt-report-builder"
py -3 scripts\bootstrap_env.py
```

会在 `~/.qoderwork-runtime/python/odps-ppt-report-builder/venv/` 建虚拟环境。

**取数：**

```bash
# Mac
~/.qoderwork-runtime/python/odps-ppt-report-builder/venv/bin/python \
  scripts/build_report_bundle.py --config <config>

# Windows
& "$env:USERPROFILE\.qoderwork-runtime\python\odps-ppt-report-builder\venv\Scripts\python.exe" `
  scripts\build_report_bundle.py --config <config>
```

**生成 slide plan：**

让 AI 读 `report_context.md` 后生成 `slide_plan.generated.json`。  
格式定义：`references/slide_plan_format.md`

**渲染：**

```bash
# Mac（Windows 同理替换 python 路径）
venv/bin/python scripts/render_ppt.py \
  --config <config> --payload <payload> --plan <plan>
```

还有 `render_preview_png.py`（PNG 预览）和 `render_preview_html.py`（HTML 预览）。

## 图表类型

渲染器按模板自动选图表，不用手动指定：

| 模板 | 图表 | 原因 |
|---|---|---|
| weekly_trend | 柱状图 | 趋势对比 |
| time_slot | 柱状图 | 时段并列 |
| price_band | 甜甜圈 | 占比分布 |
| new_old_mix | 甜甜圈 | 客群占比 |
| aoi_distribution | 甜甜圈 | AOI 占比 |
| city_distribution | 横条图 | 城市多要横排 |
| comparison | 横条图 | 月度对比 |
| cluster_mix | 柱状图 | 分段对比 |

slide plan 里指定了 `chart_type` 就用指定的。

## 排查

**odpscmd 找不到** — 检查 PATH 有没有加 odpscmd 的安装目录。  
**字体不对** — 确认 `report_config.template.yaml` 里的字体名在当前系统存在。  
**Logo 没出来** — `assets/taobao_flash_sale_logo.png` 不存在就不会显示，不报错。  
**图表空的** — `report_payload.json` 里对应 query 没数据。  
**Windows 上 python 不行** — 依次试 `py -3`、`python3`、`python`，都不行就先装 Python 3.10+，装的时候勾 "Add to PATH"。

## 别踩的坑

- 结论只写数据说的，别脑补原因
- 能查 schema 就查，别猜字段
- 少安装模式下别逼人装 Python
- 用户说一句话就直接跑，别写一堆解释
- `company_private_brief.md` 的内容是保密的

## 参考

- `references/workflow.md` — 完整流程
- `references/slide_plan_format.md` — slide plan 格式
- `references/brand_vi_spec.md` — 品牌色和 Logo 规范
- `references/ppt_template_catalog.md` — 模板视觉方向
- `references/setup_strategy_zh.md` — 安装策略
