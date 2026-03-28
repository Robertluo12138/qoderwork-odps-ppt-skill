# qoderwork-odps-ppt-skill

这是我们内部给 Qoder Work 用的 skill。

作用很简单：按固定 SQL 从 ODPS 取数，再按固定格式出 PPT。  
默认优先少安装方案，尽量照顾不懂技术的同事。

## 安装

1. 下载 `odps-ppt-report-builder.zip`
2. 解压后，把 `odps-ppt-report-builder` 文件夹放到 Qoder Work 的 skill 目录
   Mac：`~/.qoderwork/skills/`
   Windows：`%USERPROFILE%\.qoderwork\skills\`
3. 打开 Qoder Work，新建一个对话
4. 第一次不要只发一句“帮我生成某某 PPT”
5. 先把 `odps-ppt-report-builder/assets/quick_start_for_colleagues.md` 里的整段内容复制给 Qoder Work
6. 等第一次检查和环境准备走完，以后再直接说“帮我生成霸王茶姬三月的 PPT”这种话就可以

## 使用前确认

- Qoder Work 已经装好
- 公司机器上的 `odpscmd` 已经能用
- 维护人已经补好 `odps-ppt-report-builder/assets/company_private_brief.md`

上面少任何一个，Qoder Work 都不该硬往下跑。  
它应该先告诉你缺什么，再按 Mac 或 Windows 分开给步骤。

## 维护说明

你只要补这一份：

- `odps-ppt-report-builder/assets/company_private_brief.md`

做法：

1. 复制 `odps-ppt-report-builder/assets/company_private_brief.template.md`
2. 改名为 `company_private_brief.md`
3. 把真实 SQL、表结构说明、PPT 写法、业务别名、保密要求补进去

这些内容如果涉及公司敏感信息，不要传公开仓库。

## 主要文件

- `odps-ppt-report-builder/SKILL.md`
- `odps-ppt-report-builder/assets/quick_start_for_colleagues.md`
- `odps-ppt-report-builder/assets/qoderwork_first_run_prompt.template.md`
- `odps-ppt-report-builder/assets/company_private_brief.template.md`

## 额外说明

- 依赖不要装进 Qoder Work 程序目录
- 如果当前机器没有 Python，先走少安装模式
- 如果机器上已经有 Python，再走脚本模式会更稳
- 这套 skill 的目标不是“炫技”，而是让运营同学第一次也能按步骤跑通
