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

function Convert-ColorSpec {
    param([string]$Color)
    if ([string]::IsNullOrWhiteSpace($Color)) { return "RGB(0,0,0)" }
    $trimmed = $Color.Trim()
    if ($trimmed -match '^RGB\s*\(') { return $trimmed }
    if ($trimmed -match '^#?([0-9A-Fa-f]{6})$') {
        $hex = $Matches[1]
        $r = [Convert]::ToInt32($hex.Substring(0, 2), 16)
        $g = [Convert]::ToInt32($hex.Substring(2, 2), 16)
        $b = [Convert]::ToInt32($hex.Substring(4, 2), 16)
        return ("RGB({0},{1},{2})" -f $r, $g, $b)
    }
    return $trimmed
}

function Get-RgbTriplet {
    param([string]$Color)
    $formula = Convert-ColorSpec $Color
    if ($formula -match 'RGB\s*\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)') {
        return @([int]$Matches[1], [int]$Matches[2], [int]$Matches[3])
    }
    return @(0, 0, 0)
}

function New-RgbFormula {
    param([int]$R, [int]$G, [int]$B)
    $rr = [Math]::Max(0, [Math]::Min(255, $R))
    $gg = [Math]::Max(0, [Math]::Min(255, $G))
    $bb = [Math]::Max(0, [Math]::Min(255, $B))
    return ("RGB({0},{1},{2})" -f $rr, $gg, $bb)
}

function Mix-ColorSpec {
    param([string]$A, [string]$B, [double]$T)
    $aa = Get-RgbTriplet $A
    $bb = Get-RgbTriplet $B
    $tt = [Math]::Max(0.0, [Math]::Min(1.0, $T))
    return New-RgbFormula `
        ([int][Math]::Round($aa[0] + ($bb[0] - $aa[0]) * $tt)) `
        ([int][Math]::Round($aa[1] + ($bb[1] - $aa[1]) * $tt)) `
        ([int][Math]::Round($aa[2] + ($bb[2] - $aa[2]) * $tt))
}

function Adjust-ColorSpec {
    param([string]$Color, [double]$Amount)
    if ($Amount -ge 0) {
        return Mix-ColorSpec $Color "RGB(255,255,255)" $Amount
    }
    return Mix-ColorSpec $Color "RGB(0,0,0)" ([Math]::Abs($Amount))
}

function Set-TextStyle {
    param($Shape, [double]$Size = 10, [string]$Color = "RGB(0,0,0)", [bool]$Bold = $false, [int]$HAlign = 1)
    Set-CellFormula $Shape "Char.Font" "FONT(""Arial"")"
    Set-CellFormula $Shape "Char.Size" ("{0} pt" -f $Size)
    Set-CellFormula $Shape "Char.Color" (Convert-ColorSpec $Color)
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
        [double]$RoundingPx = 0,
        [double]$FillTransparency = 0,
        [double]$LineTransparency = 0
    )
    $shape = $script:page.DrawRectangle((To-X $X), (To-Y ($Y + $H)), (To-X ($X + $W)), (To-Y $Y))
    Set-CellFormula $shape "LineColor" (Convert-ColorSpec $Line)
    Set-CellFormula $shape "FillForegnd" (Convert-ColorSpec $Fill)
    Set-CellFormula $shape "FillPattern" "1"
    Set-CellFormula $shape "LineWeight" ("{0} pt" -f $LineWeight)
    Set-CellFormula $shape "LinePattern" ([string]$LinePattern)
    Set-CellFormula $shape "Rounding" ("{0} in" -f (To-Len $RoundingPx))
    Set-CellFormula $shape "FillForegndTrans" ("{0}%" -f $FillTransparency)
    Set-CellFormula $shape "LineColorTrans" ("{0}%" -f $LineTransparency)
    Add-ToLayer $shape
    return $shape
}

function Add-OutlineRectPx {
    param(
        [double]$X, [double]$Y, [double]$W, [double]$H,
        [string]$Line = "RGB(0,0,0)",
        [double]$LineWeight = 1.0,
        [int]$LinePattern = 1,
        [double]$LineTransparency = 0
    )
    $shape = $script:page.DrawRectangle((To-X $X), (To-Y ($Y + $H)), (To-X ($X + $W)), (To-Y $Y))
    Set-CellFormula $shape "LineColor" (Convert-ColorSpec $Line)
    Set-CellFormula $shape "LineWeight" ("{0} pt" -f $LineWeight)
    Set-CellFormula $shape "LinePattern" ([string]$LinePattern)
    Set-CellFormula $shape "LineColorTrans" ("{0}%" -f $LineTransparency)
    Set-CellFormula $shape "FillPattern" "0"
    Add-ToLayer $shape
    return $shape
}

function Add-GradientRectPx {
    param(
        [double]$X, [double]$Y, [double]$W, [double]$H,
        [string]$StartColor,
        [string]$EndColor,
        [int]$Steps = 10,
        [string]$Direction = "vertical"
    )
    $stepsSafe = [Math]::Max(2, $Steps)
    for ($i = 0; $i -lt $stepsSafe; $i++) {
        $t = if ($stepsSafe -le 1) { 0.0 } else { $i / [double]($stepsSafe - 1) }
        $fill = Mix-ColorSpec $StartColor $EndColor $t
        if ($Direction -eq "horizontal") {
            $sx = $X + ($W * $i / $stepsSafe)
            $sw = $W / $stepsSafe
            Add-RectPx $sx $Y ($sw + 0.25) $H $fill $fill 0 0 0 | Out-Null
        } else {
            $sy = $Y + ($H * $i / $stepsSafe)
            $sh = $H / $stepsSafe
            Add-RectPx $X $sy $W ($sh + 0.25) $fill $fill 0 0 0 | Out-Null
        }
    }
}

function Add-ShadowRectPx {
    param(
        [double]$X, [double]$Y, [double]$W, [double]$H,
        [string]$Line = "RGB(0,0,0)",
        [string]$Fill = "RGB(255,255,255)",
        [double]$LineWeight = 1.0,
        [double]$OffsetX = 4,
        [double]$OffsetY = 4,
        [double]$ShadowTransparency = 55,
        [double]$RoundingPx = 0
    )
    Add-RectPx ($X + $OffsetX) ($Y + $OffsetY) $W $H "RGB(120,120,120)" "RGB(120,120,120)" 0 0 $RoundingPx $ShadowTransparency 100 | Out-Null
    return Add-RectPx $X $Y $W $H $Line $Fill $LineWeight 1 $RoundingPx
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
    Set-CellFormula $shape "FillForegndTrans" "100%"
    Set-CellFormula $shape "LineColorTrans" "100%"
    Set-TextStyle $shape $FontSize $FontColor $Bold $HAlign
    if ($Angle -ne 0) { Set-CellFormula $shape "Angle" ("{0} deg" -f $Angle) }
    Add-ToLayer $shape
    if ($null -eq $script:textShapes) { $script:textShapes = @() }
    $script:textShapes += $shape
    return $shape
}

function Add-MaskedTextPx {
    param(
        [double]$X, [double]$Y, [double]$W, [double]$H,
        [string]$Text,
        [double]$FontSize = 10,
        [string]$FontColor = "RGB(0,0,0)",
        [bool]$Bold = $false,
        [int]$HAlign = 1,
        [string]$BackFill = "RGB(255,255,255)",
        [double]$Angle = 0
    )
    Add-RectPx $X $Y $W $H $BackFill $BackFill 0 0 0 | Out-Null
    return Add-TextPx $X $Y $W $H $Text $FontSize $FontColor $Bold $HAlign $Angle
}

function Add-SegmentedBitBarPx {
    param(
        [double]$X, [double]$Y, [double]$W, [double]$H,
        [string]$Label = "S",
        [string]$Text = "",
        [double]$WhiteW = 0,
        [double]$TabW = 11,
        [string]$Line = "RGB(20,17,241)",
        [string]$TabFill = "RGB(80,80,83)",
        [string]$TextColor = "RGB(0,0,0)"
    )
    if ($WhiteW -le 0 -or $WhiteW -gt $W) { $WhiteW = $W }
    if ($WhiteW -gt $TabW) {
        Add-RectPx ($X + $TabW) $Y ($WhiteW - $TabW) $H "RGB(255,255,255)" "RGB(255,255,255)" 0 0 0 | Out-Null
    }
    $outline = Add-RectPx $X $Y $W $H $Line "RGB(255,255,255)" 1.1 1 0
    Set-CellFormula $outline "FillPattern" "0"
    Add-RectPx $X $Y $TabW $H $Line $TabFill 0.8 | Out-Null
    Add-TextPx ($X + 1) ($Y + 1) ([Math]::Max(2, $TabW - 2)) ($H - 2) $Label 6 "RGB(255,255,255)" $true 1 | Out-Null
    if ($Text -ne "") {
        $textW = [Math]::Max(12, $WhiteW - $TabW - 4)
        Add-TextPx ($X + $TabW + 2) ($Y + 1) $textW ($H - 2) $Text 7 $TextColor $true 1 | Out-Null
    }
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
    Set-CellFormula $shape "LineColor" (Convert-ColorSpec $Line)
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
        [int]$Pattern = 1,
        [double]$FillTransparency = 0,
        [double]$LineTransparency = 0
    )
    [double[]]$visioPoints = New-Object double[] ($Points.Count)
    for ($i = 0; $i -lt $Points.Count; $i += 2) {
        $visioPoints[$i] = To-X $Points[$i]
        $visioPoints[$i + 1] = To-Y $Points[$i + 1]
    }
    $shape = $script:page.DrawPolyline($visioPoints, 0)
    Set-CellFormula $shape "LineColor" (Convert-ColorSpec $Line)
    Set-CellFormula $shape "FillForegnd" (Convert-ColorSpec $Fill)
    Set-CellFormula $shape "FillPattern" "1"
    Set-CellFormula $shape "LineWeight" ("{0} pt" -f $Weight)
    Set-CellFormula $shape "LinePattern" ([string]$Pattern)
    Set-CellFormula $shape "FillForegndTrans" ("{0}%" -f $FillTransparency)
    Set-CellFormula $shape "LineColorTrans" ("{0}%" -f $LineTransparency)
    Add-ToLayer $shape
    return $shape
}

function Add-DepthCapPx {
    param(
        [double]$X, [double]$Y, [double]$W, [double]$H,
        [double]$DepthDx = 7,
        [double]$DepthDy = 5,
        [string]$Line = "RGB(0,0,0)",
        [string]$Fill = "RGB(255,255,255)",
        [string]$TopFill = "",
        [string]$RightFill = "",
        [double]$Weight = 0.8
    )
    if ([string]::IsNullOrWhiteSpace($TopFill)) { $TopFill = Adjust-ColorSpec $Fill 0.18 }
    if ([string]::IsNullOrWhiteSpace($RightFill)) { $RightFill = Adjust-ColorSpec $Fill -0.12 }
    [double[]]$top = @($X,$Y, ($X+$DepthDx),($Y-$DepthDy), ($X+$W+$DepthDx),($Y-$DepthDy), ($X+$W),$Y, $X,$Y)
    [double[]]$right = @(($X+$W),$Y, ($X+$W+$DepthDx),($Y-$DepthDy), ($X+$W+$DepthDx),($Y+$H-$DepthDy), ($X+$W),($Y+$H), ($X+$W),$Y)
    Add-PolygonPx $top $Line $TopFill $Weight 1 | Out-Null
    Add-PolygonPx $right $Line $RightFill $Weight 1 | Out-Null
}

function Add-IsometricBlockPx {
    param(
        [double]$X, [double]$Y, [double]$W, [double]$H,
        [string]$Line = "RGB(0,0,0)",
        [string]$Fill = "RGB(255,255,255)",
        [double]$DepthDx = 7,
        [double]$DepthDy = 5,
        [double]$Weight = 1.0,
        [string]$TopFill = "",
        [string]$RightFill = ""
    )
    Add-DepthCapPx $X $Y $W $H $DepthDx $DepthDy $Line $Fill $TopFill $RightFill $Weight
    Add-RectPx $X $Y $W $H $Line $Fill $Weight 1 0 | Out-Null
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
        [bool]$Cross = $false,
        [string]$SlashDirection = "/"
    )
    Add-RectPx $X $Y $W $H "RGB(0,0,0)" (Convert-ColorSpec $Fill) 1.0 1 0 | Out-Null
    if ($SlashColor -ne "") {
        if ($SlashDirection -eq "\") {
            Add-LinePx $X $Y ($X+$W) ($Y+$H) $SlashColor 0.85 1 $false $false 1 | Out-Null
        } else {
            Add-LinePx $X ($Y+$H) ($X+$W) $Y $SlashColor 0.85 1 $false $false 1 | Out-Null
        }
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
        [hashtable]$Palette = @{},
        [double]$GapX = 0,
        [double]$GapY = 0
    )
    $defaults = @{
        cyan = @{ fill = "RGB(92,211,239)"; slash = ""; cross = $false; slashDir = "/" }
        white = @{ fill = "RGB(255,255,255)"; slash = ""; cross = $false; slashDir = "/" }
        blue_slash = @{ fill = "RGB(255,255,255)"; slash = "RGB(64,190,230)"; cross = $false; slashDir = "/" }
        red_slash = @{ fill = "RGB(255,255,255)"; slash = "RGB(255,0,0)"; cross = $false; slashDir = "\" }
        cross = @{ fill = "RGB(255,255,255)"; slash = ""; cross = $true; slashDir = "/" }
        red = @{ fill = "RGB(255,0,0)"; slash = ""; cross = $false; slashDir = "/" }
    }
    foreach ($key in $Palette.Keys) { $defaults[$key] = $Palette[$key] }
    for ($r=0; $r -lt $Matrix.Count; $r++) {
        for ($c=0; $c -lt $Matrix[$r].Count; $c++) {
            $token = $Matrix[$r][$c]
            if (-not $defaults.ContainsKey($token)) { $token = "white" }
            $style = $defaults[$token]
            $sx = $X + $c * ($CellW + $GapX)
            $sy = $Y + $r * ($CellH + $GapY)
            $slashDir = if ($style.ContainsKey("slashDir")) { $style["slashDir"] } else { "/" }
            Add-CellStyledPx $sx $sy $CellW $CellH $style["fill"] $style["slash"] $style["cross"] $slashDir
        }
    }
}

function Add-SeparatedStackedMatrixPx {
    param(
        [double]$X, [double]$Y,
        [string[][]]$Matrix,
        [double]$CellW = 13,
        [double]$CellH = 17,
        [double]$GapX = 3,
        [double]$GapY = 3,
        [double]$DepthDx = 6,
        [double]$DepthDy = 4,
        [int]$DepthLayers = 1,
        [hashtable]$Palette = @{}
    )
    $defaults = @{
        cyan = @{ fill = "RGB(92,211,239)"; slash = ""; cross = $false; slashDir = "/" }
        white = @{ fill = "RGB(255,255,255)"; slash = ""; cross = $false; slashDir = "/" }
        blue_slash = @{ fill = "RGB(255,255,255)"; slash = "RGB(64,190,230)"; cross = $false; slashDir = "/" }
        red_slash = @{ fill = "RGB(255,255,255)"; slash = "RGB(255,0,0)"; cross = $false; slashDir = "\" }
        cross = @{ fill = "RGB(255,255,255)"; slash = ""; cross = $true; slashDir = "/" }
        red = @{ fill = "RGB(255,0,0)"; slash = ""; cross = $false; slashDir = "/" }
    }
    foreach ($key in $Palette.Keys) { $defaults[$key] = $Palette[$key] }

    for ($layer=$DepthLayers; $layer -ge 1; $layer--) {
        $dx = $DepthDx * $layer
        $dy = $DepthDy * $layer
        for ($r=0; $r -lt $Matrix.Count; $r++) {
            for ($c=0; $c -lt $Matrix[$r].Count; $c++) {
                $token = $Matrix[$r][$c]
                if (-not $defaults.ContainsKey($token)) { $token = "white" }
                $style = $defaults[$token]
                $sx = $X + $c * ($CellW + $GapX)
                $sy = $Y + $r * ($CellH + $GapY)
                Add-DepthCapPx $sx $sy $CellW $CellH $dx $dy "RGB(0,0,0)" $style["fill"] "" "" 0.65
            }
        }
    }

    Add-CellMatrixPx $X $Y $Matrix $CellW $CellH $Palette $GapX $GapY
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

if (-not [string]::IsNullOrWhiteSpace($ReferenceImage)) {
    throw "Disabled: ReferenceImage import is forbidden. Keep the source image outside Visio and draw only native Visio shapes."
}

Remove-LayerShapes $script:page $LayerName
$script:drawLayer = $null
try { $script:drawLayer = $script:page.Layers.ItemU($LayerName) } catch { $script:drawLayer = $script:page.Layers.Add($LayerName) }
$script:textShapes = @()

# Add project-specific drawing calls below. Draw order should be:
# shadows/backgrounds -> gradients -> depth caps -> front faces/cell matrices -> thin lines -> thick arrows -> text.
# Example:
# $C = @{ module_blue = "#DDEAF7"; line = "#000000"; cyan = "#5CD3EF"; red = "#E8332A" }
# Add-ShadowRectPx 20 20 120 80 $C.line $C.module_blue 1.5 4 4 55 10 | Out-Null
# Add-GradientRectPx 20 20 120 18 "#F7FBFF" $C.module_blue 8 "vertical"
# Add-FilledHArrowPx 150 60 220 8 20 16 "RGB(0,0,0)"
# Add-TextPx 20 45 120 25 "Module" 16 "RGB(0,0,0)" $true 1 | Out-Null
# Stacked-grid example:
# $matrix = @(
#   @("blue_slash","blue_slash","cyan","red_slash"),
#   @("cyan","cyan","cyan","red_slash")
# )
# $palette = @{ cyan = @{ fill = $C.cyan; slash = ""; cross = $false; slashDir = "/" }; red_slash = @{ fill = "#FFFFFF"; slash = $C.red; cross = $false; slashDir = "\" } }
# Add-SeparatedStackedMatrixPx 40 100 $matrix 13 17 3 3 6 4 1 $palette

foreach ($textShape in $script:textShapes) {
    try { $textShape.BringToFront() } catch {}
}

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
