param(
    [string]$TargetPath = "",
    [string]$PageName = "VisioCopy_Trace",
    [string]$ReferenceImage = "",
    [string]$LayerName = "ManualRedraw_HiRes",
    [double]$SourceWidthPx = 1118.0,
    [double]$SourceHeightPx = 909.0,
    [double]$Scale = 0.01,
    [switch]$DryRun
)

$ErrorActionPreference = "Stop"

$PageWidth = $SourceWidthPx * $Scale
$PageHeight = $SourceHeightPx * $Scale

function To-X { param([double]$Px) return $Px * $Scale }
function To-Y { param([double]$Py) return $PageHeight - ($Py * $Scale) }
function To-Len { param([double]$Px) return $Px * $Scale }

function Set-CellFormula {
    param($Shape, [string]$Cell, [string]$Formula)
    try { $Shape.CellsU($Cell).FormulaU = $Formula } catch {}
}

function Set-TextStyle {
    param($Shape, [double]$Size = 10, [string]$Color = "RGB(0,0,0)", [bool]$Bold = $false, [int]$HAlign = 1)
    Set-CellFormula $Shape "Char.Font" "FONT(""Arial"")"
    Set-CellFormula $Shape "Char.Size" ("{0} pt" -f $Size)
    Set-CellFormula $Shape "Char.Color" $Color
    Set-CellFormula $Shape "Para.HorzAlign" ([string]$HAlign)
    Set-CellFormula $Shape "VerticalAlign" "1"
    Set-CellFormula $Shape "TextBlock.LeftMargin" "0 pt"
    Set-CellFormula $Shape "TextBlock.RightMargin" "0 pt"
    Set-CellFormula $Shape "TextBlock.TopMargin" "0 pt"
    Set-CellFormula $Shape "TextBlock.BottomMargin" "0 pt"
    if ($Bold) { Set-CellFormula $Shape "Char.Style" "1" } else { Set-CellFormula $Shape "Char.Style" "0" }
}

function Add-ToLayer {
    param($Shape)
    if ($script:drawLayer -ne $null) {
        try { $script:drawLayer.Add($Shape, 1) } catch {}
    }
}

function Add-RectPx {
    param(
        [double]$X, [double]$Y, [double]$W, [double]$H,
        [string]$Line = "RGB(0,0,0)",
        [string]$Fill = "RGB(255,255,255)",
        [double]$LineWeight = 1.0,
        [int]$LinePattern = 1,
        [double]$RoundingPx = 0
    )
    $shape = $script:page.DrawRectangle((To-X $X), (To-Y ($Y + $H)), (To-X ($X + $W)), (To-Y $Y))
    Set-CellFormula $shape "LineColor" $Line
    Set-CellFormula $shape "FillForegnd" $Fill
    Set-CellFormula $shape "FillPattern" "1"
    Set-CellFormula $shape "LineWeight" ("{0} pt" -f $LineWeight)
    Set-CellFormula $shape "LinePattern" ([string]$LinePattern)
    Set-CellFormula $shape "Rounding" ("{0} in" -f (To-Len $RoundingPx))
    Add-ToLayer $shape
    return $shape
}

function Add-OutlineRectPx {
    param(
        [double]$X, [double]$Y, [double]$W, [double]$H,
        [string]$Line = "RGB(0,0,0)",
        [double]$LineWeight = 1.0,
        [int]$LinePattern = 1
    )
    $shape = $script:page.DrawRectangle((To-X $X), (To-Y ($Y + $H)), (To-X ($X + $W)), (To-Y $Y))
    Set-CellFormula $shape "LineColor" $Line
    Set-CellFormula $shape "LineWeight" ("{0} pt" -f $LineWeight)
    Set-CellFormula $shape "LinePattern" ([string]$LinePattern)
    Set-CellFormula $shape "FillPattern" "0"
    Add-ToLayer $shape
    return $shape
}

function Add-TextPx {
    param(
        [double]$X, [double]$Y, [double]$W, [double]$H,
        [string]$Text,
        [double]$FontSize = 10,
        [string]$FontColor = "RGB(0,0,0)",
        [bool]$Bold = $false,
        [int]$HAlign = 1,
        [double]$Angle = 0
    )
    $shape = $script:page.DrawRectangle((To-X $X), (To-Y ($Y + $H)), (To-X ($X + $W)), (To-Y $Y))
    $shape.Text = $Text
    Set-CellFormula $shape "LinePattern" "0"
    Set-CellFormula $shape "FillPattern" "0"
    Set-TextStyle $shape $FontSize $FontColor $Bold $HAlign
    if ($Angle -ne 0) { Set-CellFormula $shape "Angle" ("{0} deg" -f $Angle) }
    Add-ToLayer $shape
    return $shape
}

function Add-LinePx {
    param(
        [double]$X1, [double]$Y1, [double]$X2, [double]$Y2,
        [string]$Line = "RGB(0,0,0)",
        [double]$Weight = 1.0,
        [int]$Pattern = 1,
        [bool]$EndArrow = $true,
        [bool]$BeginArrow = $false,
        [int]$ArrowSize = 2
    )
    $shape = $script:page.DrawLine((To-X $X1), (To-Y $Y1), (To-X $X2), (To-Y $Y2))
    Set-CellFormula $shape "LineColor" $Line
    Set-CellFormula $shape "LineWeight" ("{0} pt" -f $Weight)
    Set-CellFormula $shape "LinePattern" ([string]$Pattern)
    if ($EndArrow) {
        Set-CellFormula $shape "EndArrow" "5"
        Set-CellFormula $shape "EndArrowSize" ([string]$ArrowSize)
    }
    if ($BeginArrow) {
        Set-CellFormula $shape "BeginArrow" "5"
        Set-CellFormula $shape "BeginArrowSize" ([string]$ArrowSize)
    }
    Add-ToLayer $shape
    return $shape
}

function Add-PolygonPx {
    param(
        [double[]]$Points,
        [string]$Line = "RGB(0,0,0)",
        [string]$Fill = "RGB(255,255,255)",
        [double]$Weight = 1.0,
        [int]$Pattern = 1
    )
    [double[]]$visioPoints = New-Object double[] ($Points.Count)
    for ($i = 0; $i -lt $Points.Count; $i += 2) {
        $visioPoints[$i] = To-X $Points[$i]
        $visioPoints[$i + 1] = To-Y $Points[$i + 1]
    }
    $shape = $script:page.DrawPolyline($visioPoints, 0)
    Set-CellFormula $shape "LineColor" $Line
    Set-CellFormula $shape "FillForegnd" $Fill
    Set-CellFormula $shape "FillPattern" "1"
    Set-CellFormula $shape "LineWeight" ("{0} pt" -f $Weight)
    Set-CellFormula $shape "LinePattern" ([string]$Pattern)
    Add-ToLayer $shape
    return $shape
}

function Add-FilledHArrowPx {
    param(
        [double]$X1, [double]$Yc, [double]$X2,
        [double]$Shaft = 8,
        [double]$Head = 24,
        [double]$HeadH = 18,
        [string]$Color = "RGB(0,0,0)"
    )
    if ($X2 -ge $X1) {
        [double[]]$pts = @($X1,($Yc-$Shaft/2), ($X2-$Head),($Yc-$Shaft/2), ($X2-$Head),($Yc-$HeadH), $X2,$Yc, ($X2-$Head),($Yc+$HeadH), ($X2-$Head),($Yc+$Shaft/2), $X1,($Yc+$Shaft/2), $X1,($Yc-$Shaft/2))
    } else {
        [double[]]$pts = @($X1,($Yc-$Shaft/2), ($X2+$Head),($Yc-$Shaft/2), ($X2+$Head),($Yc-$HeadH), $X2,$Yc, ($X2+$Head),($Yc+$HeadH), ($X2+$Head),($Yc+$Shaft/2), $X1,($Yc+$Shaft/2), $X1,($Yc-$Shaft/2))
    }
    Add-PolygonPx $pts $Color $Color 0 1 | Out-Null
}

function Add-CellStyledPx {
    param(
        [double]$X, [double]$Y, [double]$W, [double]$H,
        [string]$Fill,
        [string]$SlashColor = "",
        [bool]$Cross = $false
    )
    Add-RectPx $X $Y $W $H "RGB(0,0,0)" $Fill 1.0 1 0 | Out-Null
    if ($SlashColor -ne "") {
        Add-LinePx $X ($Y+$H) ($X+$W) $Y $SlashColor 0.85 1 $false $false 1 | Out-Null
    }
    if ($Cross) {
        Add-LinePx $X $Y ($X+$W) ($Y+$H) "RGB(0,0,0)" 0.75 1 $false $false 1 | Out-Null
        Add-LinePx $X ($Y+$H) ($X+$W) $Y "RGB(0,0,0)" 0.75 1 $false $false 1 | Out-Null
    }
}

function Add-VisibleStackOutlinePx {
    param(
        [double]$X, [double]$Y,
        [double]$W, [double]$H,
        [int]$Layers = 5,
        [double]$Dx = 5,
        [double]$Dy = 4,
        [string]$Line = "RGB(0,0,0)"
    )
    for ($l=$Layers; $l -ge 1; $l--) {
        $ox = $l*$Dx
        $oy = -$l*$Dy
        Add-OutlineRectPx ($X+$ox) ($Y+$oy) $W $H $Line 0.75 1 | Out-Null
        Add-LinePx ($X+$ox) ($Y+$oy) ($X+$ox-$Dx) ($Y+$oy+$Dy) $Line 0.7 1 $false $false 1 | Out-Null
        Add-LinePx ($X+$W+$ox) ($Y+$oy) ($X+$W+$ox-$Dx) ($Y+$oy+$Dy) $Line 0.7 1 $false $false 1 | Out-Null
    }
}

function Add-CellMatrixPx {
    param(
        [double]$X, [double]$Y,
        [string[][]]$Matrix,
        [double]$CellW = 13,
        [double]$CellH = 17,
        [hashtable]$Palette = @{}
    )
    $defaults = @{
        cyan = @{ fill = "RGB(92,211,239)"; slash = ""; cross = $false }
        white = @{ fill = "RGB(255,255,255)"; slash = ""; cross = $false }
        blue_slash = @{ fill = "RGB(255,255,255)"; slash = "RGB(64,190,230)"; cross = $false }
        red_slash = @{ fill = "RGB(255,255,255)"; slash = "RGB(255,0,0)"; cross = $false }
        cross = @{ fill = "RGB(255,255,255)"; slash = ""; cross = $true }
        red = @{ fill = "RGB(255,0,0)"; slash = ""; cross = $false }
    }
    foreach ($key in $Palette.Keys) { $defaults[$key] = $Palette[$key] }
    for ($r=0; $r -lt $Matrix.Count; $r++) {
        for ($c=0; $c -lt $Matrix[$r].Count; $c++) {
            $style = $defaults[$Matrix[$r][$c]]
            Add-CellStyledPx ($X+$c*$CellW) ($Y+$r*$CellH) $CellW $CellH $style.fill $style.slash $style.cross
        }
    }
}

function Remove-LayerShapes {
    param($Page, [string]$Name)
    $toDelete = @()
    foreach ($shape in $Page.Shapes) {
        try {
            if ($shape.LayerCount -gt 0) {
                for ($i = 1; $i -le $shape.LayerCount; $i++) {
                    if ($shape.Layer($i).NameU -eq $Name) {
                        $toDelete += $shape
                        break
                    }
                }
            }
        } catch {}
    }
    foreach ($shape in $toDelete) { try { $shape.Delete() } catch {} }
}

if ($DryRun) {
    [pscustomobject]@{
        PageName = $PageName
        SourceWidthPx = $SourceWidthPx
        SourceHeightPx = $SourceHeightPx
        Scale = $Scale
        PageWidthIn = $PageWidth
        PageHeightIn = $PageHeight
    } | ConvertTo-Json
    exit 0
}

if ([string]::IsNullOrWhiteSpace($TargetPath)) {
    throw "Pass -TargetPath or copy this scaffold into a project-specific redraw script."
}

$TargetPath = [System.IO.Path]::GetFullPath($TargetPath)
$targetDir = [System.IO.Path]::GetDirectoryName($TargetPath)
$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$backup = Join-Path $targetDir ("visio-copy.before-redraw.{0}.vsdx" -f $timestamp)
Copy-Item -LiteralPath $TargetPath -Destination $backup -Force

try { $visio = [Runtime.InteropServices.Marshal]::GetActiveObject("Visio.Application") } catch { $visio = New-Object -ComObject Visio.Application }
$visio.Visible = $true
try { $visio.UserControl = $true } catch {}

$doc = $null
foreach ($d in $visio.Documents) {
    if ([string]::Equals($d.FullName, $TargetPath, [System.StringComparison]::OrdinalIgnoreCase)) { $doc = $d; break }
}
if ($null -eq $doc) { $doc = $visio.Documents.Open($TargetPath) }

$script:page = $null
try { $script:page = $doc.Pages.ItemU($PageName) } catch {
    $script:page = $doc.Pages.Add()
    $script:page.NameU = $PageName
    $script:page.Name = $PageName
}

$script:page.PageSheet.CellsU("PageWidth").FormulaU = ("{0} in" -f $PageWidth)
$script:page.PageSheet.CellsU("PageHeight").FormulaU = ("{0} in" -f $PageHeight)

if (-not [string]::IsNullOrWhiteSpace($ReferenceImage) -and (Test-Path -LiteralPath $ReferenceImage)) {
    $baseLayer = $null
    try { $baseLayer = $script:page.Layers.ItemU("TraceBase_HiRes") } catch { $baseLayer = $script:page.Layers.Add("TraceBase_HiRes") }
    $base = $script:page.Import([System.IO.Path]::GetFullPath($ReferenceImage))
    $base.NameU = "locked_trace_base"
    $base.CellsU("PinX").FormulaU = ("{0} in" -f ($PageWidth / 2.0))
    $base.CellsU("PinY").FormulaU = ("{0} in" -f ($PageHeight / 2.0))
    $base.CellsU("Width").FormulaU = ("{0} in" -f $PageWidth)
    $base.CellsU("Height").FormulaU = ("{0} in" -f $PageHeight)
    $base.CellsU("LocPinX").FormulaU = "Width*0.5"
    $base.CellsU("LocPinY").FormulaU = "Height*0.5"
    foreach ($cell in @("LockMove","LockWidth","LockHeight","LockDelete","LockRotate","LockAspect")) {
        Set-CellFormula $base $cell "1"
    }
    try { $baseLayer.Add($base, 1) } catch {}
    try { $base.SendToBack() } catch {}
}

Remove-LayerShapes $script:page $LayerName
$script:drawLayer = $null
try { $script:drawLayer = $script:page.Layers.ItemU($LayerName) } catch { $script:drawLayer = $script:page.Layers.Add($LayerName) }

# Add project-specific drawing calls below. Draw order should be:
# backgrounds -> modules -> stack outlines -> front cell matrices -> thin lines -> thick arrows -> text.
# Example:
# Add-RectPx 20 20 120 80 "RGB(0,0,0)" "RGB(224,234,246)" 1.5 1 10 | Out-Null
# Add-FilledHArrowPx 150 60 220 8 20 16 "RGB(0,0,0)"
# Add-TextPx 20 45 120 25 "Module" 16 "RGB(0,0,0)" $true 1 | Out-Null
# Stacked-grid example:
# $matrix = @(
#   @("blue_slash","blue_slash","cyan","red_slash"),
#   @("cyan","cyan","cyan","red_slash")
# )
# Add-VisibleStackOutlinePx 40 100 52 34 4 5 4
# Add-CellMatrixPx 40 100 $matrix 13 17

$doc.Save()
$preview = Join-Path $targetDir ("visio-copy.preview.{0}.png" -f $timestamp)
$script:page.Export($preview)

[pscustomobject]@{
    PageName = $PageName
    LayerName = $LayerName
    Backup = $backup
    Preview = $preview
    Mapping = ("{0}x{1} px source, 1 px = {2} in" -f $SourceWidthPx, $SourceHeightPx, $Scale)
} | ConvertTo-Json
