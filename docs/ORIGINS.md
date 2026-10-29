# 来源与复用说明

本仓库的专用 `ecnu-blockchain-report` 技能根据一次实际获用户认可的周报修改流程重新编写。它保留可重复的方法及案例，不是OpenAI原插件的打包副本，也不要求安装同一插件缓存。

## 上游技能来源

上次11.13 PPT修订实际使用的主技能为 `presentations:Presentations`；这次整理可复用技能时使用 `skill-creator`。模板跟随、视觉检查和最终验收方法参考以下本机文件。路径记录当时的来源位置，不是仓库运行依赖。

| 来源 | 本机审阅路径 | SHA256 |
|---|---|---|
| Presentations | `$USERPROFILE/.codex/plugins/cache/openai-primary-runtime/presentations/26.1011.51657/skills/presentations/SKILL.md` | `aff4d98d964c625697ced28c0ad142a5a3c06c59f434596755216dfa5213e8f4` |
| template_following | `$USERPROFILE/.codex/plugins/cache/openai-primary-runtime/presentations/26.1011.51657/skills/presentations/references/template_following.md` | `881875960ffa6bc59c9f6351552d9f499ac053884c780f0725cc8ccf1125c4d2` |
| finalization | `$USERPROFILE/.codex/plugins/cache/openai-primary-runtime/presentations/26.1011.51657/skills/presentations/references/finalization.md` | `32f8ff54baba353425f5b5ca11f0390f4daa991630a5bd2f2a2727f2c1726dab` |
| style_guidelines | `$USERPROFILE/.codex/plugins/cache/openai-primary-runtime/presentations/26.1011.51657/skills/presentations/style_guidelines.md` | `b682d1beb638b6ce77ecab86f83afcbde7a68207a749bc7994d8fd805feec619` |
| skill-creator | `$USERPROFILE/.codex/skills/.system/skill-creator/SKILL.md` | `cccd291077ec57c6f50ca6529f0f3fb93212da09473effb2fcec808e81b21288` |

上游Presentations文档还包含与其宿主环境、会话API和专有运行时相关的指令。本仓库按独立PPTX模板与脚本使用方式重新描述流程，没有复制那些接口或声称它们在普通Python环境中可用。

在本机检查的插件版本目录及其相邻根目录中，未发现名为LICENSE、NOTICE或COPYING的独立文件。这仅是检查记录，不构成许可判断。本仓库不打包上游插件源码、技能全文或私有运行时；保留技能名称与校验值作为来源记录。

## 实践来源

已获认可案例来自 `Minjie2003/hotrevoke` 的提交 `dfbf5372bced57ee97a58598548cfe9502b00b7c`：

- `presentations/11_13_week9_paper_revision/HotRevoke_Week9_查询契约与证据边界.pptx`：12页成品，正文7页。
- `readable_deck.py` 与 `report_content.json`：原生图形、表格、连线、四张实验图及讲稿的生成方式。
- `SPEAKER_NOTES.md`：连续例子、计量边界和常见问答。
- `validation.json` 与 `chart_validation.json`：实际渲染、编辑性和数据核对范围。

[原始公开提交](https://github.com/Minjie2003/hotrevoke/commit/dfbf5372bced57ee97a58598548cfe9502b00b7c)

案例保留原有研究语境。演示如何使用模板，不表示后续研究也得到相同结果。模板中的HotRevoke数据若用于示例，应标明其来源和时间；新课题使用自己的证据。

## 模板与素材

学校／学院的标志和模板视觉元素来自用户提供并要求保留的参考PPT。学校标志、字体和其他第三方素材的权利归各自权利人；本仓库的整理与脚本许可不自动覆盖这些素材。仓库用于汇报模板参考，不表示学校或实验室官方发布、认证或背书。

发布版本通过素材清单列出实际附带的参考PPT路径与SHA256：`templates/reference/approved-week9.pptx`为获认可成品，`templates/reference/ecnu-progress-original.pptx`为原参考模板，`templates/ecnu-report-starter.pptx`为可供填写的起点。字体使用本机已安装字体，不随仓库分发字体文件。参考PPT内部若有其他作者内容，应保留来源标识，通用示例与原始案例分开命名。

本机来源路径已将用户名规范化为`$USERPROFILE`；SHA256按当时读取的原文件计算。
