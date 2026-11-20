$ErrorActionPreference = 'Stop'

function Open-PptSource {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Application, [Parameter(Mandatory)][string]$Path)

    $file = Get-Item -LiteralPath $Path
    if ($file.PSIsContainer -or $file.Extension -ine '.pptx') {
        throw "Expected a PowerPoint .pptx file: $Path"
    }
    foreach ($open in $Application.Presentations) {
        if ($open.FullName -ieq $file.FullName) { return $open }
    }
    # ReadOnly, Untitled, WithWindow
    return $Application.Presentations.Open($file.FullName, -1, 0, -1)
}

function New-PptWorkingCopy {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Application, [Parameter(Mandatory)]$Source,
        [Parameter(Mandatory)][string]$OutputPath)

    $dest = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)
    if ([IO.Path]::GetExtension($dest) -ine '.pptx') {
        throw "Output must be a .pptx file: $dest"
    }
    if ($dest -ieq $Source.FullName -or (Test-Path -LiteralPath $dest)) {
        throw "Choose an unused output path: $dest"
    }
    $parent = Split-Path -LiteralPath $dest
    if (!(Test-Path -LiteralPath $parent -PathType Container)) {
        throw "Create the output directory first: $parent"
    }
    # ppSaveAsOpenXMLPresentation = 24
    $Source.SaveCopyAs($dest, 24)
    return $Application.Presentations.Open($dest, 0, 0, -1)
}

function Get-PptShapeById {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Slide, [Parameter(Mandatory)][int]$Id)

    function Find-ChildShape($items) {
        for ($i = 1; $i -le $items.Count; $i++) {
            $item = $items.Item($i)
            if ($item.Id -eq $Id) { return $item }
            if ($item.Type -eq 6) {
                $match = Find-ChildShape $item.GroupItems
                if ($null -ne $match) { return $match }
            }
        }
    }
    $shape = Find-ChildShape $Slide.Shapes
    if ($null -eq $shape) { throw "Shape.Id $Id not found on slide $($Slide.SlideIndex)." }
    return $shape
}

function Copy-PptComponent {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$SourceShape, [Parameter(Mandatory)]$DestinationSlide,
        [Parameter(Mandatory)][single]$Left, [Parameter(Mandatory)][single]$Top,
        [Parameter(Mandatory)][ValidateScript({$_ -gt 0})][single]$Width,
        [Parameter(Mandatory)][ValidateScript({$_ -gt 0})][single]$Height)

    # Components may be resized independently; avoid distorting photos.
    if ($SourceShape.Type -in 11,13 -and
        [Math]::Abs((($Width / $Height) / ($SourceShape.Width / $SourceShape.Height)) - 1) -gt 0.005) {
        throw 'Picture aspect ratio would change; crop the photo in PowerPoint or adjust the box.'
    }

    $range = $null
    for ($attempt = 0; $attempt -lt 4; $attempt++) {
        try {
            $SourceShape.Copy()
            $range = $DestinationSlide.Shapes.Paste()
            break
        } catch {
            # Clipboard contention on an interactive desktop is usually transient.
            if ($attempt -eq 3) { throw }
            Start-Sleep -Milliseconds (120 * ($attempt + 1))
        }
    }
    if ($range.Count -ne 1) { throw 'Expected one pasted component. Inspect this slide before retrying.' }
    $shape = $range.Item(1)
    try {
        $locked = $shape.LockAspectRatio
        $shape.LockAspectRatio = 0
        $shape.Left = $Left
        $shape.Top = $Top
        $shape.Width = $Width
        $shape.Height = $Height
        $shape.LockAspectRatio = $locked
    } catch {
        $shape.Delete()
        throw
    }
    return $shape
}

Export-ModuleMember -Function Open-PptSource,New-PptWorkingCopy,Get-PptShapeById,Copy-PptComponent
