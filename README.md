# qoderwork-odps-ppt-skill

这个仓库放的是我们自己用的 Qoder Work skill。

它主要做一件事：按固定 SQL 从 ODPS 取数，再按固定格式整理成 PPT。

默认思路是尽量用 PowerPoint 原生对象来出内容，后面改字、改表、改图会方便一些。

如果你只是拿来装给同事用，直接看 `odps-ppt-report-builder.zip` 就行。  
如果你是维护这套 skill 的人，再看 `odps-ppt-report-builder/` 里的文件。

## 同事怎么用

1. 下载 `odps-ppt-report-builder.zip`
2. 解压后，把 `odps-ppt-report-builder` 文件夹放到 Qoder Work 的 skill 目录
3. Qoder Work 的 skill 目录：
   Mac：`~/.qoderwork/skills/`
   Windows：`%USERPROFILE%\.qoderwork\skills\`
4. 打开 Qoder Work，新开一个对话
5. 把 `odps-ppt-report-builder/assets/qoderwork_first_run_prompt.template.md` 里的内容发给 Qoder Work
6. 按提示补日期、业务范围和其他参数，等它出结果

## 使用前先确认

- 电脑上已经装好 Qoder Work
- 公司机器已经配置好 `odpscmd`
- 维护人已经补好 `odps-ppt-report-builder/assets/company_private_brief.md`

如果机器上刚好有 Python，可以直接走脚本模式，结果会更稳一点。  
如果没有 Python，也可以先按少安装模式用。

## 维护人要补什么

1. 复制 `odps-ppt-report-builder/assets/company_private_brief.template.md`
2. 改名为 `odps-ppt-report-builder/assets/company_private_brief.md`
3. 把真实 SQL、表结构说明、PPT 写法补进去

这些内容如果涉及公司敏感信息，就不要传到公开仓库。

## 怎么看仓库是不是完整的

首页能看到下面这些，就说明仓库结构没问题：

- `odps-ppt-report-builder/`
- `odps-ppt-report-builder.zip`
- `README.md`

点进 `odps-ppt-report-builder/` 之后，继续确认有：

- `SKILL.md`
- `assets/`
- `references/`
- `scripts/`

这些都在的话，这个仓库就是完整可用的版本。

## 额外提醒

- 依赖不要装进 Qoder Work 程序目录
- 如果当前机器不方便直接生成 `.pptx`，先让 Qoder Work 产出完整讲稿和预览文件也可以
