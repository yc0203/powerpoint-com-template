---
name: powerpoint-com-template
description: 用 Windows 桌面版 Microsoft PowerPoint 的 COM 接口制作、改版或核查 PPTX，制作前向作者收取内容文档与图表素材并确认一套模板，先编排再用 COM 逐页制作、对照复核；亦可维护本地模板库。启用本 Skill 后，PowerPoint 制作和保存必须由 COM 完成。
---

# 桌面 PowerPoint 模板制作

**强制：用本机 Windows 桌面版 Microsoft PowerPoint 的 COM 对象 'PowerPoint.Application' 完成演示文稿的打开、复制、编辑、保存和逐页导出。** 用户没有明确再说 COM，也照此执行。MCP、Spaces、PowerPoint 网页版、LibreOffice、PptxGenJS、python-pptx 和直接改 PPTX/XML 均不能作为此 Skill 的制作路径。可以用其他工具准备照片、文字或做只读校验；将素材放进演示文稿仍由 COM 执行。COM 不可用或当前执行主机不是有桌面会话的 Windows 时，说明缺少哪项条件，停止制作。不得悄悄更换途径。

## 安装和首次使用

- 前提：当前 Codex 任务在本机 Windows 运行；已安装并能正常打开桌面版 PowerPoint；可运行 PowerShell 和调用 Office COM。无须配置 MCP。需要可见过程时，Windows 桌面会话必须处于可交互状态。
- 分享或迁移时，把整个 'powerpoint-com-template' 文件夹复制到接收者 Codex 的个人 'skills' 目录（设置了 'CODEX_HOME' 时为其下的 'skills'）。须保留 'SKILL.md'、'agents'、'assets'、'references' 和 'scripts'；只复制 Markdown 会丢模板和脚本。新对话中输入 '$powerpoint-com-template'。
- 可直接说：「用 $powerpoint-com-template，在本机 PowerPoint 做一份 15 页的技术汇报，先问我用哪套模板，并请我提供内容文档及图表素材。技术路线可编辑，让我看见过程。免费试用的自制示例模板编号为 starter-research。」也可以直接指定模板编号或路径，作为对选型问题的回答。机器相关路径从当前会话取得；不依赖制作者的 E 盘、账户名、历史输出或第三方 PowerShell 模块。

## 按模板做新稿或改版

**先确认一套模板，再收齐足以制作的内容。** 按 [问答与制作提示词](references/authoring-prompt.md) 一次问清缺项：作者选哪套，以及是否已提供内容文档和对应插图、图表数据。文档和素材已给就不重复问；作者已给完整正文或明确要求基于其素材共同起草，可以直接使用。模板未选定或关键内容缺失时，只读整理和核对，待作者回答后再创建或改写 PPT；不把参考模板的旧内容当成新稿材料。仅检查 PPT 或登记模板不触发制作问答。

1. **先整理章节和逐页图文计划。** 根据作者文档建立章节顺序、每页核心结论和素材出处；把图片、说明、表格数据对应到具体页。依 [内容编排与图表规格](references/content-planning.md) 为每个对象定位置、宽高、对齐、图注和边框，缺失或冲突的证据记清楚。未要求审稿时，计划完成就继续制作。
2. **从选定的一套模板找原生组件。** 核对背景、字体、母版、文本框、图框、箭头、构图和页面布局；给每页标注参考页及组件。按 [单模板选页指南](references/layout-selection.md) 选择不同正文版式；**相邻两页不要用完全相同的正文布局**，但统一页眉、页脚、配色和字号层级。同一术语统一称呼与强调方式，关键术语和关键结论句统一用同一红色，详见图表规格。
3. **只用桌面 PowerPoint COM 操作副本。** 新稿以模板副本起稿；改稿以原稿副本起稿，再从同一套选定模板复制原生对象。使用绝对路径匹配窗口；作者要求观看时，用 'View.GotoSlide' 逐页展示。详见 [COM 手册](references/com-workflow.md)；使用随包示例时，读 [样式档案](references/starter-research.md)。
4. **与模板并排复核，再修改。** 通过 PowerPoint 导出成稿，按逐页计划找到模板对应页，比较版心、构图、字体、背景和组件；同时复核图表来源、尺寸、对齐、线型、红色强调及相邻页版式。发现问题用 COM 修正，重导受影响页面，保存并重开验证；只报告实际通过的项目。

## 维护模板库

用户要求收录自己的模板时，按 [模板库清单](references/templates.md) 新增独立的 PPTX 副本、代表页预览和样式档案。使用同一 COM 工作流；模板库是此 Skill 随包携带的本地资产。
