# 部署策略说明

这个文档专门给团队内部部署时参考，目标是让 `Qoder Work` 在 `Mac` 和 `Windows` 上都更稳。

## 结论先说

不要把依赖装进 Qoder Work 的程序目录。

推荐分三层：

1. Skill 文件层

- 放到 `~/.qoderwork/skills/odps-ppt-report-builder/`
- 这里只放 `SKILL.md`、模板、脚本、参考文档

2. 用户级共享运行时

- `Mac`:
  - `~/.qoderwork-runtime/python/odps-ppt-report-builder/`
- `Windows`:
  - `%USERPROFILE%\.qoderwork-runtime\python\odps-ppt-report-builder\`
- 这里放虚拟环境、Python 包、缓存
- 好处是：
  - 不污染 Qoder Work 安装目录
  - 不需要管理员权限
  - 同一台电脑后续新 skill 也可以复用这套 Python 运行时思路

3. 机器级工具

- 这类工具应该单独装在系统或用户环境里：
  - Python
  - odpscmd
  - PowerPoint 或 WPS
  - Node.js
  - git
- 这些不应该被塞进 skill 文件夹

## 为什么不建议直接全局 `pip install`

- 容易和别的项目冲突
- Windows 上经常会出现 PATH、权限、版本不一致问题
- 后面 skill 一多，很难知道哪个包是谁装的

## 为什么也不建议放进 Qoder Work 程序目录

- 程序升级后容易丢
- 不利于多个 skill 复用
- 有些公司电脑对程序目录写权限更严格

## 当前这个 skill 的推荐策略

### 对普通运营同学

- 默认走零开发环境模式
- 如果机器已经有：
  - `odpscmd`
  - Qoder Work
- 那就可以先完成：
  - 取数
  - 结构整理
  - 生成 slide manuscript
- 如果 `.pptx` 渲染受阻，也不要卡死

### 对已经有 Python 的同学

- 再启用脚本模式
- 用 `scripts/bootstrap_env.py` 初始化共享运行时
- 用脚本做：
  - 规范化取数结果
  - 生成 slide plan
  - 渲染样例 PPT

## Mac 和 Windows 的重点差异

### Mac

- 命令通常是 `python3`
- 路径分隔符是 `/`
- 常见问题是：
  - Python 已安装但 `pip` 不完整
  - 权限弹窗

### Windows

- 优先尝试 `py -3`
- 路径分隔符是 `\`
- 常见问题是：
  - `python` 命令不可用但 `py` 可用
  - PATH 配置混乱
  - PowerShell 执行策略
  - 用户目录包含空格

## 给 Qoder Work 的写法建议

- 一次只给用户一个命令
- 先判断系统，再给命令
- 避免一句话里混很多前提
- 对 Windows 命令单独列出来
- 如果失败，要让 AI 先解释“失败在哪一步”，再给下一步

## Node.js 要不要装全局

如果未来某个 skill 真需要 Node.js，建议装成机器级工具，而不是装进 skill 文件夹。

原因：

- Node.js 本身就是一个运行时
- 后续多个 skill 可能共用
- 安装到程序目录没有复用价值

但当前这个 `ODPS -> PPT` skill，不要主动引入 Node.js，除非业务上确实需要。
