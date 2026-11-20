[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$Path,
    [int[]]$Pages,
    [string]$OutputDir,
    [ValidateRange(400,3840)][int]$PixelWidth = 1600,
    [switch]$Show
)
$ErrorActionPreference = 'Stop'

function Read-Shape($shape) {
    $hasText = $shape.HasTextFrame -eq -1 -and $shape.TextFrame2.HasText -eq -1
    $textValue = if ($hasText) { $shape.TextFrame2.TextRange.Text } else { $null }
    $kind = [int]$shape.Type
    $geom = if ($kind -eq 1) { [int]$shape.AutoShapeType } else { $null }
    $nested = @()
    if ($kind -eq 6) {
        $nested = @(for ($k = 1; $k -le $shape.GroupItems.Count; $k++) {
            Read-Shape ($shape.GroupItems.Item($k))
        })
    }
    [ordered]@{
        id = [int]$shape.Id
        name = [string]$shape.Name
        type = $kind
        autoShapeType = $geom
        left = [Math]::Round($shape.Left, 1)
        top = [Math]::Round($shape.Top, 1)
        width = [Math]::Round($shape.Width, 1)
        height = [Math]::Round($shape.Height, 1)
        text = $textValue
        children = $nested
    }
}

$src = Get-Item -LiteralPath $Path
if ($src.PSIsContainer -or $src.Extension -ine '.pptx') { throw 'Expected a .pptx file.' }
$app = New-Object -ComObject PowerPoint.Application
$deck = $null
$ownsDeck = $false
try {
    $deck = @($app.Presentations | Where-Object { $_.FullName -ieq $src.FullName })[0]
    if (!$deck) {
        $withWindow = if ($Show) { -1 } else { 0 }
        $deck = $app.Presentations.Open($src.FullName, -1, 0, $withWindow)
        $ownsDeck = $true
    }
    $slideCount = [int]$deck.Slides.Count
    if ($slideCount -lt 1) { throw 'No slides to inspect.' }
    if (!$Pages) { $Pages = @(1..$slideCount) }
    $pageList = @($Pages | Select-Object -Unique)
    foreach ($i in $pageList) {
        if ($i -lt 1 -or $i -gt $slideCount) { throw "Slide $i outside 1..$slideCount." }
    }
    $exportDir = $null
    if ($OutputDir) {
        $exportDir = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputDir)
        if (!(Test-Path -LiteralPath $exportDir -PathType Container)) {
            $null = [IO.Directory]::CreateDirectory($exportDir)
        }
        foreach ($i in $pageList) {
            $candidate = Join-Path $exportDir ('slide-{0:000}.png' -f $i)
            if (Test-Path -LiteralPath $candidate) { throw "Preview already exists: $candidate" }
        }
    }
    $window = $null
    if ($Show -and $deck.Windows.Count -gt 0) {
        $app.Visible = -1
        $window = $deck.Windows.Item(1)
        $window.Activate()
    }
    $w = [double]$deck.PageSetup.SlideWidth
    $h = [double]$deck.PageSetup.SlideHeight
    $rows = @(foreach ($i in $pageList) {
        $slide = $deck.Slides.Item($i)
        if ($window) { $window.View.GotoSlide($i) }
        $png = $null
        if ($exportDir) {
            $png = Join-Path $exportDir ('slide-{0:000}.png' -f $i)
            $slide.Export($png, 'PNG', $PixelWidth, [int][Math]::Round($PixelWidth * $h / $w))
        }
        [ordered]@{
            index = $i
            name = [string]$slide.Name
            preview = $png
            shapes = @(for ($j = 1; $j -le $slide.Shapes.Count; $j++) { Read-Shape ($slide.Shapes.Item($j)) })
        }
    })
    [ordered]@{
        path = $src.FullName
        slideCount = $slideCount
        widthPoints = $w
        heightPoints = $h
        slides = $rows
    } | ConvertTo-Json -Depth 64
}
finally {
    if ($ownsDeck -and $deck -and !$Show) { $deck.Close() }
}
