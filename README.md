# qoderwork-odps-ppt-skill

给 Qoder Work 用的经营分析 PPT 自动生成 skill。  
从 ODPS 拉数据，AI 做分析，一键出 PPT。

## 安装

**拖拽安装（推荐）**  
下载仓库里的 `odps-ppt-report-builder.zip`，拖进 Qoder Work 即可。

**手动安装**  
把 `odps-ppt-report-builder/` 文件夹放到 skill 目录：
- Mac: `~/.qoderwork/skills/`
- Windows: `%USERPROFILE%\.qoderwork\skills\`

## 使用前

1. 确保 `odpscmd` 已安装且能正常连接 ODPS
2. 复制 `assets/company_private_brief.template.md` → `company_private_brief.md`，补上真实 SQL 和业务口径
3. （可选）把 `assets/taobao_flash_sale_logo.png` 替换成官方 Logo

## 使用

首次使用把 `assets/quick_start_for_colleagues.md` 的内容发给 Qoder Work。  
后续直接说"帮我生成三月经营分析 PPT"就行了。

## 项目结构

```
odps-ppt-report-builder/
├── SKILL.md                    # skill 主指令
├── requirements.txt            # Python 依赖
├── assets/                     # 配置 & 素材
│   ├── report_config.template.yaml
│   ├── company_private_brief.template.md
│   ├── quick_start_for_colleagues.md
│   └── taobao_flash_sale_logo.png
├── scripts/                    # 核心脚本
│   ├── bootstrap_env.py        # 环境初始化
│   ├── build_report_bundle.py  # ODPS 取数
│   ├── fill_slide_plan.py      # 生成 slide plan
│   ├── render_ppt.py           # 渲染 PPTX
│   ├── render_preview_png.py   # PNG 预览
│   └── render_preview_html.py  # HTML 预览
└── references/                 # 格式定义 & 规范文档
```

## 跨平台

Mac 和 Windows 都支持。字体、路径、Python 命令的差异已在代码和配置里处理好：

| | Mac | Windows |
|---|---|---|
| 中文字体 | Hiragino Sans GB | Microsoft YaHei |
| Python | `python3` | `py -3` |
| venv 路径 | `venv/bin/python` | `venv\Scripts\python.exe` |

## 品牌色

品牌色 `#FF6200`，辅助色 `#FFA020`。详见 `references/brand_vi_spec.md`。

## 注意

- `company_private_brief.md` 含公司敏感信息，已加入 `.gitignore`，不要提交
- Logo 占位图需替换为官方素材后再出正式报告
- 没有 Python 的电脑走少安装模式，有 Python 走脚本模式
