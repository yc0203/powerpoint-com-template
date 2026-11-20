# PowerPoint COM 模板制作

简体中文 · [English](README.en.md)

这是一个 Codex Skill，用 **Windows 桌面版 Microsoft PowerPoint 的 COM 接口**按模板制作或改版可编辑的 .pptx。每份稿只用一套模板，制作前先让作者选模板、提供内容文档与图表素材，再编排章节和逐页版式。

需要本地 Windows 任务、已安装并可正常打开的桌面版 Microsoft PowerPoint 和 PowerShell。无需 MCP、Python 依赖或 Office 加载项。网页版、远程 Linux 或只会聊天的 AI 无法执行桌面 COM。

## 这段话直接复制给 AI

```text
请在我本机的 Windows Codex 环境中，从
https://github.com/yc0203/powerpoint-com-template
下载这个仓库。先阅读 README、powerpoint-com-template/SKILL.md，
再把整个 powerpoint-com-template 子文件夹安装到本机 Codex 的个人 Skills
目录；保留其中的 scripts、references、agents 和 assets。优先遵循本机配置好的
Skills 路径；已有同名版本先告诉我，不要覆盖，也不要更改全局配置。
如果仓库需要登录，使用我已有的 GitHub 授权，不要问我要令牌。
安装后告诉我 SKILL.md 的绝对路径以及如何用 $powerpoint-com-template 调用。
```

安装后，在新对话里复制：

```text
用 $powerpoint-com-template 在本机桌面 PowerPoint 制作一份可编辑的 PPT。
开始前先问我用哪一套模板，并请我提供含正文、插图、图注和图表数据的内容文档。
先按章节整理，逐页规划图表的位置、尺寸、对齐和边框；只用选定的一套模板。
统一重点术语和关键结论的红色，相邻正文页不要完全同构。
用 COM 制作并让我看见逐页操作。完稿与模板对应页对照，修改后重导并重开核对。
```

## 安装和模板

1. 通过 Git 下载仓库或用 GitHub 的 **Code → Download ZIP** 获取。
2. 将仓库内的 [powerpoint-com-template/](powerpoint-com-template/) 整个复制到 Codex 用户技能目录。例如，标准用户目录可用 ~/.agents/skills/，自定义安装以本机识别路径为准。安装后应存在 `.../skills/powerpoint-com-template/SKILL.md`。
3. 新建 Codex 对话，调用 `$powerpoint-com-template`。没有显示时，可把本机 SKILL.md 的绝对路径告诉 Codex，并确认所用任务运行在本机 Windows。

[自制入门模板](powerpoint-com-template/assets/starter-research/reference.pptx)包含封面、目录、证据页、四步路线、原生表格和计划页；编号为 `starter-research`。也可以指定自己有权使用的本地 PPTX。仓库不附第三方模板。

制作时，Skill 用选中的模板安排图片、表格和章节，尽量保留原生文本框、分区、连接线和表格。完成后用 PowerPoint 导图，与参考页并排检查，再修正和重开。新建原生公式是否能完整写入取决于 Office 版本和 COM；不能时会说明。原稿和模板不直接修改。

[查看完整工作流](powerpoint-com-template/SKILL.md) · [查看模板清单](powerpoint-com-template/references/templates.md)
