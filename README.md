# Qoder Work ODPS 自动汇报 Skill

给运营和业务同学用的 Qoder Work skill。

它的目标很简单：按固定 SQL 从 ODPS 取数，整理固定表结构，并按约定模板生成 PPT 报告。

## 你只需要关心什么

- `odps-ppt-report-builder/`：skill 本体
- `odps-ppt-report-builder.zip`：给同事分发时直接下载的压缩包

## 同事怎么用

1. 下载仓库里的 `odps-ppt-report-builder.zip`
2. 解压后，把 `odps-ppt-report-builder` 文件夹放到 Qoder Work 的 skill 目录
3. 打开 Qoder Work，新开一个对话
4. 先阅读 `odps-ppt-report-builder/assets/qoderwork_first_run_prompt.template.md`
5. 把里面的首轮提示词发给 Qoder Work
6. 按提示补充日期、业务范围和需求，等待生成报告

Qoder Work skill 目录：

- Mac：`~/.qoderwork/skills/`
- Windows：`%USERPROFILE%\.qoderwork\skills\`

## 首次使用前

- 电脑上已经安装 Qoder Work
- 公司电脑已经按内部规范配置好 `odpscmd`
- 维护人已经补齐内部私有说明文件 `assets/company_private_brief.md`

说明：

- 如果电脑有 Python，skill 会优先走更稳定的脚本流程
- 如果电脑没有 Python，也可以先按零配置模式使用

## 维护人要做的事

1. 复制 `odps-ppt-report-builder/assets/company_private_brief.template.md`
2. 重命名为 `odps-ppt-report-builder/assets/company_private_brief.md`
3. 填入公司内部真实 SQL、表结构说明、PPT 编写规则

不要把真实 SQL、敏感表名、字段口径提交到公开仓库。

## 怎么检查仓库是不是完整的

打开 GitHub 仓库首页后，确认能看到这些内容：

- `odps-ppt-report-builder/`
- `odps-ppt-report-builder.zip`
- `README.md`

点进 `odps-ppt-report-builder/` 后，继续确认有：

- `SKILL.md`
- `assets/`
- `references/`
- `scripts/`

如果这些都在，说明这个 skill 仓库已经是完整可用版本。

## 常见提醒

- 不要把依赖装进 Qoder Work 程序安装目录
- 不要把公司敏感取数逻辑提交到公开仓库
- 如果生成 `.pptx` 受限，先让 Qoder Work 输出结构化讲稿和预览文件
