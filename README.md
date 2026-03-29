# qoderwork-odps-ppt-skill

面向 Qoder Work 的 PPT 自动生成 Skill。  
从 ODPS 取数 → AI 分析 → 自动出品牌经营诊断 PPT，颜色和 Logo 严格遵循淘宝闪购品牌 VI 标准。

## 一、安装

### 第 1 步：下载 skill

下载本仓库的 zip 包，解压得到 `odps-ppt-report-builder` 文件夹。

### 第 2 步：放到 Qoder Work 的 skill 目录

**macOS：**

```bash
cp -r odps-ppt-report-builder ~/.qoderwork/skills/odps-ppt-report-builder
```

**Windows（在 PowerShell 或文件管理器中操作）：**

```powershell
Copy-Item -Recurse odps-ppt-report-builder "$env:USERPROFILE\.qoderwork\skills\odps-ppt-report-builder"
```

> 如果 `~/.qoderwork/skills/` 目录不存在，手动创建即可。

### 第 3 步：替换 Logo（可选但推荐）

`assets/taobao_flash_sale_logo.png` 当前是 AI 生成的占位图。  
请找品牌组要官方**透明底 PNG Logo**，替换这个文件。文件名保持不变。

### 第 4 步：补充私有业务说明

1. 复制 `assets/company_private_brief.template.md`
2. 改名为 `assets/company_private_brief.md`
3. 把真实的 SQL、表结构说明、业务别名、PPT 写法要求写进去

> ⚠️ `company_private_brief.md` 包含公司敏感信息，**不要上传到公开仓库**。

## 二、使用前确认

在任何电脑上首次使用前，确保以下 3 项全部满足：

| 检查项 | macOS 检查命令 | Windows 检查命令 |
|---|---|---|
| Qoder Work 已安装 | 打开应用确认 | 打开应用确认 |
| `odpscmd` 可用 | `which odpscmd` | `where odpscmd` |
| 私有说明文件存在 | `ls assets/company_private_brief.md` | `dir assets\company_private_brief.md` |

上面少任何一个，Qoder Work 都不该硬往下跑。它应该先提示缺什么，再按 Mac 或 Windows 分开给步骤。

## 三、使用方法

### 首次使用

打开 Qoder Work，新建一个对话，把 `assets/quick_start_for_colleagues.md` 里的**整段内容**复制粘贴发送。  
Qoder Work 会自动：

1. 检查 skill 位置
2. 检查 `odpscmd` 是否可用
3. 检查私有说明文件
4. 如果机器有 Python，自动初始化环境
5. 取数 + 分析 + 生成 PPT

### 后续使用

直接说自然语言即可：

- "帮我生成霸王茶姬三月的经营分析 PPT"
- "帮我重跑上周周报"
- "帮我出本月月报"

## 四、项目结构

```
odps-ppt-report-builder/
├── SKILL.md                          # Qoder Work 读的 skill 指令主文件
├── requirements.txt                  # Python 依赖（PyYAML, python-pptx）
│
├── assets/                           # 配置和素材
│   ├── report_config.template.yaml   # 报告配置模板（颜色、字体、slide定义）
│   ├── company_private_brief.template.md  # 业务说明模板
│   ├── quick_start_for_colleagues.md      # 首次使用 prompt
│   ├── qoderwork_first_run_prompt.template.md
│   └── taobao_flash_sale_logo.png    # 品牌 Logo（请替换为官方版）
│
├── scripts/                          # 核心脚本
│   ├── common.py                     # 公共工具函数
│   ├── bootstrap_env.py              # 初始化 Python 虚拟环境
│   ├── build_report_bundle.py        # ODPS 取数 + 中间包
│   ├── fill_slide_plan.py            # 填充 slide plan
│   ├── render_ppt.py                 # 渲染 PPTX（主力渲染器）
│   ├── render_preview_png.py         # 渲染 PNG 预览（contact sheet）
│   ├── render_preview_html.py        # 渲染 HTML 预览
│   └── install_to_qoderwork.py       # 安装 skill 到 Qoder Work
│
├── references/                       # 参考文档
│   ├── brand_vi_spec.md              # 淘宝闪购品牌 VI 标准色和 Logo 规范
│   ├── ppt_template_catalog.md       # PPT 模板类型和视觉方向
│   ├── slide_plan_format.md          # slide plan JSON 格式定义
│   ├── setup_strategy_zh.md          # 安装策略说明
│   └── workflow.md                   # 完整工作流
│
└── output/                           # 运行时产出（不需要提交到仓库）
```

## 五、跨平台说明

本 skill 同时支持 macOS 和 Windows。以下差异已在代码中处理：

| 差异点 | macOS | Windows |
|---|---|---|
| Python 命令 | `python3` | `py -3` 或 `python` |
| 中文字体 | Hiragino Sans GB | Microsoft YaHei |
| 西文字体 | Aptos | Aptos |
| 虚拟环境 Python 路径 | `venv/bin/python` | `venv\Scripts\python.exe` |
| 路径分隔符 | `/` | `\`（代码用 `pathlib` 自动处理）|

## 六、品牌 VI 标准

| 角色 | 色值 | 占比 |
|---|---|---|
| 品牌色 | `#FF6200` | 90% |
| 辅助色 | `#FFA020` | 10% |

详细规范见 `references/brand_vi_spec.md`。

## 七、维护说明

日常维护只需关注一个文件：`assets/company_private_brief.md`

如需调整颜色、字体或图表配色，修改 `assets/report_config.template.yaml` 中的 `theme` 部分。

## 八、注意事项

- 依赖不要装进 Qoder Work 程序目录
- 没有 Python 的电脑先走少安装模式，有 Python 再走脚本模式
- `company_private_brief.md` 是保密文件，不要提交到公开仓库
- Logo 文件请务必替换为官方素材后再生成正式报告
