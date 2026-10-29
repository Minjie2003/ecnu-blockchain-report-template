---
name: ecnu-blockchain-report
description: Create or revise Chinese research progress PPTX reports using the ECNU blockchain laboratory reference style, with a coherent research story, editable diagrams and charts, timed notes, and rendered-slide review. Use when the user selects this template or asks to follow its weekly reporting style.
metadata:
  short-description: 华师大区块链实验室周报模板与制作流程
---

# 华师大区块链实验室汇报

将研究进展整理成听众能跟上的汇报，并保留参考PPT的视觉结构。先读用户指定的模板和现有内容；仓库中的HotRevoke成品用于理解效果，不把它的研究结论、姓名、日期或实验编号带入新课题。

## 开始前

- 仓库中以仓库根为工作目录；通过安装器安装后，以技能根为工作目录。两种模式都使用同层的 `scripts/`、`templates/` 与 `examples/`。阅读 `scripts/build_report.py --help`、`scripts/render_report.py --help` 和 `scripts/validate_report.py --help` 的实际参数后再运行，不依赖某台机器的绝对路径。
- 确认是修改现有PPT，还是制作新一期。修改时保留原目录与文件名，先备份；新一期使用清楚的日期或周次目录，不制造多个含义不明的“最终版”。
- 以用户本次指定的模板和要求为准。查看实际参考页，核对画幅、标志、标题、字体、分隔线和页脚；不要仅凭模板名称重建外观。
- 阅读相关论文、上期汇报和原始实验说明，区分本周新增内容、沿用的已有证据和下一步计划。参考文件缺失时标出缺口，不把计划写成成果。
- 完整工作步骤见[制作与验收流程](references/workflow.md)。开始排版前读取[文字与视觉约定](references/style-guide.md)。

通用起点是 `templates/ecnu-report-starter.pptx` 与 `examples/starter.json`。`templates/reference/approved-week9.pptx` 和 `examples/hotrevoke-week9/` 展示获认可的实际成品；原参考模板在 `templates/reference/ecnu-progress-original.pptx`。选择最贴合用户任务的来源，不修改冻结参考文件。

## 讲清楚一条主线

先写一句“这次汇报希望听众理解什么”，据此安排正文。通常从具体问题进入，说明已有工作解决了什么、还需解释什么，再讲本方案和证据。让一个连续例子贯穿问题、架构与算法；新符号在首次出现时解释。

保留学校周报常用的封面、主要任务、研究进展、本周学习体会与反思、下周科研计划及感谢页。主要任务可按短期、中长期、长期组织。页数由内容和用户要求决定；参考成品的12页不是固定上限或配额。每周优先呈现新增内容，只保留理解它所需的背景。

使用浅显的中文句子，说明“谁做了什么，为什么，这一步改变什么”。减少脱离上下文的术语、小标题和碎片。正文放关键解释、图与证据；进一步推导、计量口径和常见问答放入逐页备注。

## 原生内容与证据

- 文字、表格及主要流程图尽量保持原生可编辑；连线连接到模块。图中的方向、先后和职责必须与论文或代码一致，不能为排版改变含义。
- 数值图使用来源明确的数据，保留单位、分母、条件、时间边界与误差含义。已有图表的工作簿和公式应保留，禁止靠修改缓存造出新的结果。
- 收益和成本一起解释。保留无收益、失败或未测范围，不把小规模原型的观察扩成生产容量、普遍优势或安全证明。
- 演示数据要明确标成演示。复制案例图时不得把HotRevoke的真实数值说成其他课题的新实验。

## 工具和验证

按当前环境选择可用工具。若存在可用的Presentations技能，可配合其导入与检查流程；仓库提供的独立脚本也可使用。不得把专用插件、Artifact Tool或Windows COM预设为所有机器都有。运行前读仓库README的实际依赖和命令。

编辑完成后保存PPTX，导出全部页预览并逐张查看。检查换行、遮挡、裁切、字号、连线、图例和证据口径；对照原模板检查风格。自动检查只覆盖它实际执行的项目，不能代替视觉检查。

在临时副本上抽查文字编辑、表格或图表编辑，以及移动模块后连线是否跟随。若导出PDF，它应来自同一份最终PPTX。记录文件哈希、工具实际名称、检查结果及未执行项；不能把WPS检查写成已经在微软PowerPoint检查。

科学内容复杂时，可安排独立的内容审查和数据审查。交付前整合发现的问题，再渲染受影响页面。用户请求制作PPT不自动意味着运行新实验、发邮件或推送仓库；在用户已授权的发布范围内完成对应动作。

交付一个明确的可编辑PPT入口、对应PDF（若已请求或沿用既有交付方式）和逐页讲稿。简述仍会影响使用的限制，不把临时构建文件当成正式成果。
