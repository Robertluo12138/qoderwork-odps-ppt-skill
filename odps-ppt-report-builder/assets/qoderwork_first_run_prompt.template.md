请用 `odps-ppt-report-builder` 这个 skill 帮我做这次报告，按下面的方式来：

- 默认按 Qoder Work 处理，不要按 Qoder IDE 来理解
- 先判断这台机器能不能少安装直接完成；只有确实需要时再走 Python 方案
- 除非必要，不要让我额外安装 Python、git、Node.js 或别的包管理工具
- 依赖不要装到 Qoder Work 程序目录里
- 如果需要 Python，就装到共享运行时：
  - macOS：`~/.qoderwork-runtime/`
  - Windows：`%USERPROFILE%\.qoderwork-runtime\`
- 先读 `assets/company_private_brief.md`
- 如果没有这个文件，就提醒我从 `assets/company_private_brief.template.md` 复制一份再补内容
- 取数时只用私有说明里写明的 SQL 或命令，不要自己改
- Windows 请一步一步写命令，不要省略前提
- 如果 macOS 和 Windows 的命令不一样，请分开写清楚
- 中间文件统一放在当前工作目录下的一个任务文件夹里
- 至少生成这些文件：
  - `raw_schema.txt`
  - `report_context.md`
  - `slide_plan.generated.json`
- 如果能直接生成 PPT，请优先用原生文本框、表格、图表和形状，不要整页截图，也不要把核心内容做成大贴图
- 除非我明确要求插图，否则正文不要依赖 AI 位图图片来承载关键信息
- PPT 结构按私有说明来，不要自己改大纲
- 最终结果里不要带敏感字段或隐藏口径
- 如果这台机器没法直接生成 `.pptx`，就先给完整讲稿，不要停在半路
- 如果机器上有 Python，优先用 skill 自带脚本，结果会更稳
- 如果我只是想先看效果，顺手再给一个 HTML 预览或 PNG 预览

最后请给我：

- 最终文件路径
- 3 到 5 条重点结论
- 哪些步骤卡住了，或者还缺什么数据
