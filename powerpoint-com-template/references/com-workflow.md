# Windows PowerPoint COM 操作手册

## 先检查参考稿

打开 PowerPoint，查看参考页的图片、分区和图注样式。检查脚本传入真实路径，可枚举全稿或指定页，也可导出 PNG：

```powershell
$skill = '此 Skill 文件夹的绝对路径'
$ref = '已下载或作者提供的本地 PPTX 绝对路径'
& (Join-Path $skill 'scripts\Inspect-PowerPoint.ps1') -Path $ref -Pages 1,2,3,4,5,6 -OutputDir '本任务可写的新预览目录'
```

不传 'OutputDir' 就只返回 JSON；不传 'Pages' 就检查全稿；传 '-Show' 会显示并在参考页间切换。形状的 'Id' 不是 'Shapes.Item()' 序号。组内文字到 'GroupItems' 找；更换参考文件版本应重新点查 ID。涉及启动 GUI 或受限目录时走 Codex 的权限流程。

## 打开副本并控制桌面编辑窗口

```powershell
Import-Module (Join-Path $skill 'scripts\PowerPoint-Com.psm1')
$app = New-Object -ComObject PowerPoint.Application
$app.Visible = -1
$inputDeck = Open-PptSource $app $ref
$outDeck = New-PptWorkingCopy $app $inputDeck $outPath
$win = $outDeck.Windows.Item(1)
$win.Activate()
$win.View.GotoSlide(3)
$slide = $outDeck.Slides.Item(3)
```

'ref' 是模板；'outPath' 是用户输出目录中尚不存在的绝对 '.pptx' 路径。已有内容改版时，传原内容稿创建副本，再另开参考。PowerPoint 会话可能已经打开多个文稿：按 'FullName' 匹配，避免对 'ActivePresentation' 盲写；不要退出整个应用，不关闭用户原来已打开的文稿，不保存源文件。有未保存用户编辑或多份同名源需要区分时，先明确处理对象。'SaveCopyAs(..., 24)' 保留原稿，按同一对象的 'CustomLayout' 和可见样式核对复制结果。

当用户要观看，每批几页切换可见窗口并报页号，可留短暂停留让画面更新；不要从后台一次替换整份文件。避免同时运行两条复制任务，它们共享 Windows 剪贴板。完成后保持成稿可见，停在用户关心的页面。

## 整页和原生组件怎么复用

- **整页：** 例如可用 'outDeck.Slides.InsertFromFile($ref, 1, 2, 2)' 复制模板第 2 页到输出第 1 页后。核对返回计数、页序、主题和底图；只在确认后删除对应旧页。新稿已是同一文件时，可用 '.Duplicate()' 后 '.MoveTo()'。删除对象前记录目标对象 ID，避免用固定高度批量删掉新稿的原生公式或备注。
- **组件：** 下例按清单中的页码和形状 ID 复制图题并挪到目标。覆盖普通标签的文本即可；复杂文字分段使用 'TextRange2.Characters(start, length)' 保留字体层级。复制保留填充、描边、形状、渐变和效果；跨主题粘贴之后仍要用预览检查是否被重映射。脚本针对剪贴板忙作有限重试。先用检查脚本给出 $sourcePage 和 $shapeId，再执行以下示例。

```powershell
$swatch = Get-PptShapeById -Slide ($inputDeck.Slides.Item($sourcePage)) -Id $shapeId
$label = Copy-PptComponent -SourceShape $swatch -DestinationSlide $slide -Left 49 -Top 109 -Width 320 -Height 31
$label.TextFrame2.TextRange.Text = '证据图题'
```

- **照片与分区：** 框、照片、图注是三个独立对象。先放外框，再摆照片，最后放图注；加框后如果挡图，按对象调整 'ZOrder' 或先复制底层。照片保持原有高宽比，按内容裁剪而非拉伸，图片插入用 COM 'Shapes.AddPicture' 并嵌入文件。生成图注注明概念图；真实样机与来源分开标。
- **流程和表格：** 学样式，按新内容用原生形状和 'Shapes.AddConnector' 搭图。用 'BeginConnect' / 'EndConnect' 锚定；矩形连接点可先验 4→2 为右到左，3→1 为下到上，其他几何不保证。移动节点验证箭头跟随。数据表优先保留原生表格或使用 'Shapes.AddTable'。无真实数据时，把表做成比较方案/指标矩阵，别填样例数字。
- **原生公式：** 现有公式移位置和尺寸，保留数学对象；不要给含数学的整个 'TextRange' 赋新字串。新增试用 PowerPoint 自身命令：激活目标普通视图、先显示该页、选中文本框并进入文字选择，再调用 'app.CommandBars.ExecuteMso("EquationInsertNew")'。命令可用性取决于 Office 版本和当前选择。创建后须实际验证公式填入并能用公式工具编辑；若本机无法通过 COM 可靠写入公式正文，可复制已提供且适用的 PowerPoint 原生公式，或明确报告这部分受限。不要改 OOXML 来绕过本 Skill 的 COM 制作要求。

## 逐页验收并保存

- 草稿用 COM 'Save()' 保存。用检查脚本或 'Slide.Export' 按页面比例逐页输出 PNG；全稿缩略图检查版式差异，放大检查照片、公式、图注及最密页。已修正的页重新导出。保留用户指定的输出位置。
- 核对 'TextFrame2.TextRange.BoundWidth/BoundHeight' 与形状尺寸减去文本边距；东亚标点悬挂可能产生少量误报，视觉二次确认。调整框尺寸、缩短文案或调整留白优先于一律减字体。检查底框在图文后、标题和图注不遮内容、组内文字没有移出。
- 保存稿重开后验证应可编辑的文字、表格、连线、原生公式、页序和超链接；只读解包检查 'a14:m'、'm:oMath'、'p:cxnSp' 或链接关系作为补充，不替代 PowerPoint 重开。源文件保持不变。不把上一版案例的页数、坐标、模板内容或科研成果写死到其他任务。
