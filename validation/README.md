# 本次实际验证范围

本目录保存首次公开版本的检查结果。生成新汇报后需要重新检查，新内容不会自动继承这些结果。

- `starter_structure.json`：10页、10份备注、1张原生表格、1张带嵌入工作簿的图表和2条绑定连线；关系、计数和图表数据一致。
- `starter_render.json`：从最终PPTX导出10张预览和PDF，源文件未被修改。COM注册名称为PowerPoint，但实际提供程序路径属于WPS；未据此声称微软PowerPoint测试通过。
- `starter_visual_review.json`：逐张查看10页预览，检查文字、边界、对齐、流程连线及示例数据标记。
- `editability_probe.json`：在临时副本中编辑文字、移动流程模块，确认连线跟随，保存关闭并重开后再次检查；正式源文件保持不变。
- `independent_use.json`：另一审查者用无关课题完成自定义名称安装、三页构建和结构检查，并验证安装后的文档链接。
- `builder_checks.json`、`metadata_checks.json`：不同页数、四节点流程、含小数和缺失值的数据图，以及过时元数据的拒绝检查。
- `installed_skill_structure.json`：最初安装包的10页独立构建结构检查；后续修复和最终安装复验见 `independent_use.json`。

模板示例数据仅为排版演示。实际研究案例沿用HotRevoke已发布结果，本次没有重跑其研究实验。结构检查不能替代视觉检查；模板也不能保证任意长度的后续文字都适配原版面。
