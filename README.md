---
title: 淘宝闪购品牌经营诊断报告体系
description: 依托 Qoder Work 和 ODPS 构建的数智化看板与诊断演示自动化分发系统。
---

# 淘宝闪购品牌经营诊断自动化架构 (ODPS -> PPT)

本项目专为数据/运营团队日常、周度和月度业务复盘设计。通过集成 Qoder Work AI Agent 框架与底层计算引擎 (ODPS)，打通数据提取、指标加工、多维归因分析与最终汇报级演示文稿（PPTX）的端到端自动化流程。最终生成的产物全面遵守**淘宝闪购专属视觉与品牌 VI 规范**。

---

## 一、 系统能力特性

- **数据端到端全免手动**: 直连 ODPS 跨源取数，自动化规避手动贴图与图表更新的繁冗流程。
- **智能化洞察与归因**: 深度融合大模型能力识别结构化数据表特征，自动沉淀业务核心结论与动作建议（Takeaways）。
- **专业级品牌 VI 表现**: 1:1 高保真还原运营团队手工打造的专业排版设计（全宽品牌横幅、多维渐变色阶、类别适配性智能图表引擎）。

## 二、 部署集成指南

根据使用者的技术熟练度以及所处环境，我们提供两种集成方案。

### 方式 A：一键式拖拽激活（推荐业务运营侧）

1. 从本仓库或内网分享获取最新版 `odps-ppt-report-builder.zip`。
2. 打开客户端 **Qoder Work**。
3. **直接将 zip 压缩包拖入 Qoder Work 会话窗口内**，或通过设置菜单选取 "安装 Skill" 加载该包。
4. 加载完毕后立即可通过对话激活系统功能。

### 方式 B：源码级拉取部署（推荐数据工程侧调试）

若有特殊网络环境或其他不兼容情况，可通过命令行实现静默配置：

**macOS 生态:**
```bash
cp -r odps-ppt-report-builder ~/.qoderwork/skills/odps-ppt-report-builder
```

**Windows 生态 (需 PowerShell):**
```powershell
Copy-Item -Recurse odps-ppt-report-builder "$env:USERPROFILE\.qoderwork\skills\odps-ppt-report-builder"
```
> *注：若目标节点路径 `~/.qoderwork/skills/` 不存在，请手动声明该路径目录。*

## 三、 使用前检查核验 (Pre-flight Checks)

为保证管线顺畅执行，终端节点必须通过以下关键项校验：

1. **终端算力与客户端**: 已部署 Qoder Work 客户端。
2. **底层计算探针**: 终端具备通过环境变量或相对路径执行 `odpscmd` (ODPS Client) 的能力。
3. **私有配置落位**: 项目目录树中已注入真实的业务配置（参见第四点安全合规策略）。

## 四、 安全合规与核心配置

所有的核心 SQL、表定义说明及其对应的分析重点定义，全部实行安全隔离。

1. 拷贝模板配置表：将 `assets/company_private_brief.template.md` 克隆为 `assets/company_private_brief.md`。
2. 注入真实上下文：在克隆后的文件中填充正式的数据集查询 SQL 和对应的业务定义口径。
> 🚨 **安全高压红线**：`company_private_brief.md` 以及含有机密业务逻辑的文件**严禁通过 Git 提交至公共内网代码库外流**（已纳入 `.gitignore` 过滤项）。

## 五、 项目工程结构详解

```text
odps-ppt-report-builder/
├── SKILL.md                          # [核心] AI Agent Pipeline 的调度说明文件
├── requirements.txt                  # Python 依赖清单
│
├── assets/                           # [配置层] 核心资产与参数
│   ├── report_config.template.yaml   # 渲染引擎配置文件（色彩映射、排版定义）
│   ├── company_private_brief.template.md # 业务私有定义模板
│   └── taobao_flash_sale_logo.png    # 品牌占位 Logo 素材（建议替换高清版本）
│
├── scripts/                          # [脚本层] 工具链与渲染引擎
│   ├── build_report_bundle.py        # Pipeline: 执行数据抽取并生成上下文
│   ├── fill_slide_plan.py            # Pipeline: AST及模型渲染数据视图与结论
│   ├── render_ppt.py                 # Engine: 主力 PPTX 面向对象视图渲染器
│   └── render_preview_*.py           # Engine: 旁路渲染输出轻量图片与网页看板
│
├── references/                       # [规范层] 协议与开发标准声明
│   ├── brand_vi_spec.md              # 淘宝闪购品牌 VI 大字库及色彩声明
│   ├── ppt_template_catalog.md       # PPT 模板对象封装映射说明
│   └── workflow.md                   # 跨组件业务工作流拓扑
│
└── output/                           # [输出层] 日志、运行时数据切片及成品堆栈存放位置
```

## 六、 引擎规范矩阵概览

### 跨平台适配抽象
本框架在底层抽象处理了主要操作系统的差异，开发者及运营无需侧目细节：
- 解释器寻址：自动推断 `python3` (Mac) 或 `py -3` (Win)。
- CJK字体回溯链：优先选用 `Hiragino Sans GB` 面向 Mac 系列，`Microsoft YaHei` 面向下卷兼容体系。
- 环境隔离边界：自动拉起 `venv`。

### VI 色彩表现管理与映射
- 首选品牌主色: **`#FF6200`** (承担 90% 视觉中心点)，辅助突出版式线与主卡片。
- 对比提示色: **`#FFA020`** (处理局部跳跃节点)。
- 多向图表色阶生成链: *已在底层写死并动态适配*。

---
内部维护团队联系与规范指引，请详阅代码库内 `references/` 细则文档。
