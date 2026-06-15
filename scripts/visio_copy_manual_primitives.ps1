<#
Reusable primitives for manual visio-copy redraw scripts.

Expected setup in the caller:
  $Visio = New-Object -ComObject Visio.Application
  $doc = $Visio.Documents.Add("")
  $Page = $Visio.ActivePage
  Initialize-VisioCopyCanvas -Page $Page -SourceWidthPx 1118 -SourceHeightPx 704

The helpers draw editable Visio shapes in source-image pixel coordinates.
They do not paste the source image and they do not save to redraw.vsdx.
#>

$script:VisioCopyTextJobs = @()

function Get-VisioCopyValue {
    param($Object, [string]$Name, $Default = $null)
    if ($null -eq $Object) { return $Default }
    if ($Object -is [hashtable] -and $Object.ContainsKey($Name)) { return $Object[$Name] }
    $prop = $Object.PSObject.Properties[$Name]
    if ($null -ne $prop) { return $prop.Value }
    return $Default
}

function Initialize-VisioCopyCanvas {
    param(
        [Parameter(Mandatory=$true)]$Page,
        [Parameter(Mandatory=$true)][double]$SourceWidthPx,
        [Parameter(Mandatory=$true)][double]$SourceHeightPx,
        [double]$Scale = (1.0 / 96.0),
        [string]$LayerName = "ManualRedraw_visio-copy"
    )
    $script:VisioCopyPage = $Page
    $script:VisioCopySourceWidthPx = $SourceWidthPx
    $script:VisioCopySourceHeightPx = $SourceHeightPx
    $script:VisioCopyScale = $Scale
    $script:VisioCopyPageWidth = $SourceWidthPx * $Scale
    $script:VisioCopyPageHeight = $SourceHeightPx * $Scale
    $script:VisioCopyTextJobs = @()

    $Page.PageSheet.CellsU("PageWidth").FormulaU = "$($script:VisioCopyPageWidth) in"
    $Page.PageSheet.CellsU("PageHeight").FormulaU = "$($script:VisioCopyPageHeight) in"
    $script:VisioCopyLayer = $Page.Layers.Add($LayerName)
}

function ConvertTo-VisioCopyX {
    param([double]$Px)
    return $Px * $script:VisioCopyScale
}

function ConvertTo-VisioCopyY {
    param([double]$Py)
    return $script:VisioCopyPageHeight - ($Py * $script:VisioCopyScale)
}

function ConvertTo-VisioCopyColorFormula {
    param([string]$Color)
    if ([string]::IsNullOrWhiteSpace($Color)) { return "RGB(0,0,0)" }
    if ($Color.StartsWith("#")) {
        $r = [Convert]::ToInt32($Color.Substring(1, 2), 16)
        $g = [Convert]::ToInt32($Color.Substring(3, 2), 16)
        $b = [Convert]::ToInt32($Color.Substring(5, 2), 16)
        return "RGB($r,$g,$b)"
    }
    return $Color
}

function Set-VisioCopyColorCell {
    param($Shape, [string]$Cell, [string]$Color)
    $Shape.CellsU($Cell).FormulaU = ConvertTo-VisioCopyColorFormula $Color
}

function Add-VisioCopyShapeToLayer {
    param($Shape)
    try { $script:VisioCopyLayer.Add($Shape, 1) | Out-Null } catch {}
}

function Set-VisioCopyShapeStyle {
    param(
        $Shape,
        [string]$Line = "#000000",
        [string]$Fill = "none",
        [double]$WeightPt = 0.8,
        [switch]$Dash,
        [double]$FillTransparency = 0
    )
    if ($Fill -eq "none" -or $Fill -eq "transparent") {
        $Shape.CellsU("FillPattern").FormulaU = "0"
    } else {
        $Shape.CellsU("FillPattern").FormulaU = "1"
        Set-VisioCopyColorCell $Shape "FillForegnd" $Fill
        $Shape.CellsU("FillForegndTrans").FormulaU = "$FillTransparency%"
    }

    if ($Line -eq "none" -or $Line -eq "transparent") {
        $Shape.CellsU("LinePattern").FormulaU = "0"
    } else {
        $Shape.CellsU("LinePattern").FormulaU = "1"
        Set-VisioCopyColorCell $Shape "LineColor" $Line
        $Shape.CellsU("LineWeight").FormulaU = "$WeightPt pt"
    }
    if ($Dash) { $Shape.CellsU("LinePattern").FormulaU = "2" }
    Add-VisioCopyShapeToLayer $Shape
}

function Add-VisioCopyRectPx {
    param(
        [double]$X,
        [double]$Y,
        [double]$W,
        [double]$H,
        [string]$Line = "#000000",
        [string]$Fill = "none",
        [double]$WeightPt = 0.8,
        [switch]$Dash,
        [double]$RoundPx = 0
    )
    $shape = $script:VisioCopyPage.DrawRectangle(
        (ConvertTo-VisioCopyX $X),
        (ConvertTo-VisioCopyY ($Y + $H)),
        (ConvertTo-VisioCopyX ($X + $W)),
        (ConvertTo-VisioCopyY $Y)
    )
    Set-VisioCopyShapeStyle $shape $Line $Fill $WeightPt -Dash:$Dash
    if ($RoundPx -gt 0) {
        $shape.CellsU("Rounding").FormulaU = "$(ConvertTo-VisioCopyX $RoundPx) in"
    }
    return $shape
}

function Add-VisioCopySharedGridPx {
    param(
        [double]$X,
        [double]$Y,
        [int]$Cols,
        [int]$Rows,
        [double]$CellW,
        [double]$CellH,
        [string]$Line = "#000000",
        [string]$Fill = "#FFFFFF",
        [double]$WeightPt = 0.8,
        [double]$OuterWeightPt = -1
    )
    if ($Cols -le 0 -or $Rows -le 0 -or $CellW -le 0 -or $CellH -le 0) {
        throw "invalid shared grid dimensions"
    }
    $w = $Cols * $CellW
    $h = $Rows * $CellH
    if ($OuterWeightPt -lt 0) { $OuterWeightPt = $WeightPt }
    Add-VisioCopyRectPx -X $X -Y $Y -W $w -H $h -Line $Line -Fill $Fill -WeightPt $OuterWeightPt | Out-Null
    for ($col = 1; $col -lt $Cols; $col++) {
        $x0 = $X + ($col * $CellW)
        Add-VisioCopyLinePx -X1 $x0 -Y1 $Y -X2 $x0 -Y2 ($Y + $h) -Line $Line -WeightPt $WeightPt | Out-Null
    }
    for ($row = 1; $row -lt $Rows; $row++) {
        $y0 = $Y + ($row * $CellH)
        Add-VisioCopyLinePx -X1 $X -Y1 $y0 -X2 ($X + $w) -Y2 $y0 -Line $Line -WeightPt $WeightPt | Out-Null
    }
}

function Add-VisioCopyGridCellFillPx {
    param(
        [double]$X,
        [double]$Y,
        [double]$CellW,
        [double]$CellH,
        [int]$Col,
        [int]$Row,
        [string]$Fill = "#FFFFFF"
    )
    Add-VisioCopyRectPx -X ($X + ($Col * $CellW)) -Y ($Y + ($Row * $CellH)) -W $CellW -H $CellH -Line "none" -Fill $Fill -WeightPt 0 | Out-Null
}

function Add-VisioCopyPageBackgroundPx {
    param(
        [string]$Fill = "#FFFFFF"
    )
    $shape = Add-VisioCopyRectPx 0 0 $script:VisioCopySourceWidthPx $script:VisioCopySourceHeightPx $Fill $Fill 0.1
    try { $shape.SendToBack() } catch {}
    $sentinel = "#FEFEFE"
    Add-VisioCopyRectPx 0 0 1 1 $sentinel $sentinel 0.1 | Out-Null
    Add-VisioCopyRectPx ($script:VisioCopySourceWidthPx - 1) 0 1 1 $sentinel $sentinel 0.1 | Out-Null
    Add-VisioCopyRectPx 0 ($script:VisioCopySourceHeightPx - 1) 1 1 $sentinel $sentinel 0.1 | Out-Null
    Add-VisioCopyRectPx ($script:VisioCopySourceWidthPx - 1) ($script:VisioCopySourceHeightPx - 1) 1 1 $sentinel $sentinel 0.1 | Out-Null
    return $shape
}

function Remove-VisioCopyOffPageEmptyShapes {
    param(
        [double]$ToleranceIn = 0.02
    )
    $page = $script:VisioCopyPage
    if ($null -eq $page) { return 0 }
    $pageW = [double]$page.PageSheet.CellsU("PageWidth").ResultIU
    $pageH = [double]$page.PageSheet.CellsU("PageHeight").ResultIU
    $toDelete = @()
    foreach ($shape in $page.Shapes) {
        try {
            if (($null -ne $shape.Text) -and ($shape.Text.Trim().Length -gt 0)) { continue }
            $pinX = [double]$shape.CellsU("PinX").ResultIU
            $pinY = [double]$shape.CellsU("PinY").ResultIU
            $w = [double]$shape.CellsU("Width").ResultIU
            $h = [double]$shape.CellsU("Height").ResultIU
            $left = $pinX - ($w / 2)
            $right = $pinX + ($w / 2)
            $bottom = $pinY - ($h / 2)
            $top = $pinY + ($h / 2)
            if ($left -lt (0 - $ToleranceIn) -or $right -gt ($pageW + $ToleranceIn) -or $bottom -lt (0 - $ToleranceIn) -or $top -gt ($pageH + $ToleranceIn)) {
                $toDelete += $shape
            }
        } catch {}
    }
    foreach ($shape in $toDelete) {
        try { $shape.Delete() } catch {}
    }
    return $toDelete.Count
}

function Add-VisioCopyOvalPx {
    param(
        [double]$X,
        [double]$Y,
        [double]$W,
        [double]$H,
        [string]$Line = "#000000",
        [string]$Fill = "#FFFFFF",
        [double]$WeightPt = 0.8
    )
    $shape = $script:VisioCopyPage.DrawOval(
        (ConvertTo-VisioCopyX $X),
        (ConvertTo-VisioCopyY ($Y + $H)),
        (ConvertTo-VisioCopyX ($X + $W)),
        (ConvertTo-VisioCopyY $Y)
    )
    Set-VisioCopyShapeStyle $shape $Line $Fill $WeightPt
    return $shape
}

function Add-VisioCopyDecisionDiamondPx {
    param(
        [double]$X,
        [double]$Y,
        [double]$W,
        [double]$H,
        [string]$Line = "#000000",
        [string]$Fill = "#FFFFFF",
        [double]$WeightPt = 1.4
    )
    $s = Add-VisioCopyRectPx $X $Y $W $H $Line $Fill $WeightPt
    try { $s.CellsU("Angle").FormulaU = "45 deg" } catch {}
    return $s
}

function Add-VisioCopyDocumentPx {
    param(
        [double]$X,
        [double]$Y,
        [double]$W,
        [double]$H,
        [string]$Line = "#000000",
        [string]$Fill = "#FFFFFF",
        [double]$WeightPt = 1.4,
        [double]$WavePx = 14
    )
    Add-VisioCopyRectPx $X $Y $W $H "none" $Fill 0 | Out-Null
    Add-VisioCopyPolylinePx -Points @(
        @($X,$Y),
        @(($X + $W),$Y),
        @(($X + $W),($Y + $H - $WavePx))
    ) -Line $Line -WeightPt $WeightPt | Out-Null
    $y0 = $Y + $H - $WavePx
    $pts = @(
        @(($X + $W),$y0),
        @(($X + ($W * 0.72)),($y0 + 1)),
        @(($X + ($W * 0.52)),($y0 + ($WavePx * 0.55))),
        @(($X + ($W * 0.28)),($y0 + $WavePx)),
        @($X,($y0 + ($WavePx * 0.40)))
    )
    Add-VisioCopyPolylinePx -Points $pts -Line $Line -WeightPt $WeightPt | Out-Null
    Add-VisioCopyLinePx $X ($y0 + ($WavePx * 0.40)) $X $Y $Line $WeightPt | Out-Null
}

function Add-VisioCopyCylinderPx {
    param(
        [double]$X,
        [double]$Y,
        [double]$W,
        [double]$H,
        [string]$Line = "#000000",
        [string]$Fill = "#FFFFFF",
        [double]$WeightPt = 1.4,
        [double]$CapHeightPx = 24
    )
    Add-VisioCopyRectPx $X ($Y + ($CapHeightPx / 2)) $W ($H - $CapHeightPx) "none" $Fill 0 | Out-Null
    Add-VisioCopyOvalPx $X $Y $W $CapHeightPx $Line $Fill $WeightPt | Out-Null
    Add-VisioCopyLinePx $X ($Y + ($CapHeightPx / 2)) $X ($Y + $H - ($CapHeightPx / 2)) $Line $WeightPt | Out-Null
    Add-VisioCopyLinePx ($X + $W) ($Y + ($CapHeightPx / 2)) ($X + $W) ($Y + $H - ($CapHeightPx / 2)) $Line $WeightPt | Out-Null
    Add-VisioCopyOvalPx $X ($Y + $H - $CapHeightPx) $W $CapHeightPx $Line $Fill $WeightPt | Out-Null
}

function Add-VisioCopyCutCornerRectPx {
    param(
        [double]$X,
        [double]$Y,
        [double]$W,
        [double]$H,
        [string]$Line = "#000000",
        [string]$Fill = "#FFFFFF",
        [double]$WeightPt = 1.4,
        [double]$CutX = 42,
        [double]$CutY = 42,
        [ValidateSet("top-left","top-right","bottom-left","bottom-right")][string]$Corner = "top-left"
    )
    Add-VisioCopyRectPx $X $Y $W $H "none" $Fill 0 | Out-Null
    if ($Corner -eq "top-left") {
        $pts = @(@(($X + $CutX),$Y), @(($X + $W),$Y), @(($X + $W),($Y + $H)), @($X,($Y + $H)), @($X,($Y + $CutY)))
    } elseif ($Corner -eq "top-right") {
        $pts = @(@($X,$Y), @(($X + $W - $CutX),$Y), @(($X + $W),($Y + $CutY)), @(($X + $W),($Y + $H)), @($X,($Y + $H)))
    } elseif ($Corner -eq "bottom-left") {
        $pts = @(@($X,$Y), @(($X + $W),$Y), @(($X + $W),($Y + $H)), @(($X + $CutX),($Y + $H)), @($X,($Y + $H - $CutY)))
    } else {
        $pts = @(@($X,$Y), @(($X + $W),$Y), @(($X + $W),($Y + $H - $CutY)), @(($X + $W - $CutX),($Y + $H)), @($X,($Y + $H)))
    }
    Add-VisioCopyPolylinePx -Points $pts -Line $Line -WeightPt $WeightPt -Closed | Out-Null
}

function Add-VisioCopyLinePx {
    param(
        [double]$X1,
        [double]$Y1,
        [double]$X2,
        [double]$Y2,
        [string]$Line = "#000000",
        [double]$WeightPt = 1.0,
        [switch]$Dash,
        [int]$EndArrow = 0,
        [int]$BeginArrow = 0
    )
    $shape = $script:VisioCopyPage.DrawLine(
        (ConvertTo-VisioCopyX $X1),
        (ConvertTo-VisioCopyY $Y1),
        (ConvertTo-VisioCopyX $X2),
        (ConvertTo-VisioCopyY $Y2)
    )
    Set-VisioCopyColorCell $shape "LineColor" $Line
    $shape.CellsU("LineWeight").FormulaU = "$WeightPt pt"
    if ($Dash) { $shape.CellsU("LinePattern").FormulaU = "2" }
    if ($EndArrow -ne 0) { $shape.CellsU("EndArrow").FormulaU = "$EndArrow" }
    if ($BeginArrow -ne 0) { $shape.CellsU("BeginArrow").FormulaU = "$BeginArrow" }
    Add-VisioCopyShapeToLayer $shape
    return $shape
}

function Add-VisioCopyPolylinePx {
    param(
        [Parameter(Mandatory=$true)][object[]]$Points,
        [string]$Line = "#000000",
        [double]$WeightPt = 1.0,
        [switch]$Dash,
        [int]$EndArrow = 0,
        [int]$BeginArrow = 0,
        [switch]$Closed
    )
    $shapes = @()
    if ($Points.Count -lt 2) { return $shapes }
    for ($i = 0; $i -lt ($Points.Count - 1); $i++) {
        $begin = 0
        $end = 0
        if ($i -eq 0) { $begin = $BeginArrow }
        if ($i -eq ($Points.Count - 2)) { $end = $EndArrow }
        $shapes += Add-VisioCopyLinePx `
            ([double]$Points[$i][0]) ([double]$Points[$i][1]) `
            ([double]$Points[$i + 1][0]) ([double]$Points[$i + 1][1]) `
            $Line $WeightPt -Dash:$Dash -BeginArrow $begin -EndArrow $end
    }
    if ($Closed) {
        $last = $Points[$Points.Count - 1]
        $first = $Points[0]
        $shapes += Add-VisioCopyLinePx ([double]$last[0]) ([double]$last[1]) ([double]$first[0]) ([double]$first[1]) $Line $WeightPt -Dash:$Dash
    }
    return $shapes
}

function Add-VisioCopyPolygonPx {
    param(
        [Parameter(Mandatory=$true)][object[]]$Points,
        [string]$Line = "#000000",
        [string]$Fill = "#FFFFFF",
        [double]$WeightPt = 1.0,
        [switch]$Dash,
        [double]$FillTransparency = 0
    )
    if ($Points.Count -lt 3) { return $null }

    $closed = @()
    foreach ($p in $Points) { $closed += ,$p }
    $first = $closed[0]
    $last = $closed[$closed.Count - 1]
    if (([double]$first[0] -ne [double]$last[0]) -or ([double]$first[1] -ne [double]$last[1])) {
        $closed += ,@([double]$first[0], [double]$first[1])
    }

    [double[]]$visioPoints = New-Object double[] ($closed.Count * 2)
    for ($i = 0; $i -lt $closed.Count; $i++) {
        $visioPoints[$i * 2] = ConvertTo-VisioCopyX ([double]$closed[$i][0])
        $visioPoints[($i * 2) + 1] = ConvertTo-VisioCopyY ([double]$closed[$i][1])
    }

    $shape = $script:VisioCopyPage.DrawPolyline($visioPoints, 0)
    Set-VisioCopyShapeStyle -Shape $shape -Line $Line -Fill $Fill -WeightPt $WeightPt -Dash:$Dash -FillTransparency $FillTransparency
    return $shape
}

function Add-VisioCopyTriangleSymbolPx {
    param(
        [double]$X,
        [double]$Y,
        [double]$W,
        [double]$H,
        [string]$Line = "#000000",
        [double]$WeightPt = 1.4,
        [ValidateSet("right","left","up","down")][string]$Direction = "right",
        [switch]$Bubble
    )
    if ($Direction -eq "right") {
        $pts = @(@($X,$Y), @(($X + $W),($Y + ($H / 2))), @($X,($Y + $H)))
        $bubbleX = $X + $W
        $bubbleY = $Y + ($H / 2)
    } elseif ($Direction -eq "left") {
        $pts = @(@(($X + $W),$Y), @($X,($Y + ($H / 2))), @(($X + $W),($Y + $H)))
        $bubbleX = $X
        $bubbleY = $Y + ($H / 2)
    } elseif ($Direction -eq "up") {
        $pts = @(@($X,($Y + $H)), @(($X + ($W / 2)),$Y), @(($X + $W),($Y + $H)))
        $bubbleX = $X + ($W / 2)
        $bubbleY = $Y
    } else {
        $pts = @(@($X,$Y), @(($X + ($W / 2)),($Y + $H)), @(($X + $W),$Y))
        $bubbleX = $X + ($W / 2)
        $bubbleY = $Y + $H
    }
    Add-VisioCopyPolylinePx -Points $pts -Line $Line -WeightPt $WeightPt -Closed | Out-Null
    if ($Bubble) {
        Add-VisioCopyOvalPx ($bubbleX - 5) ($bubbleY - 5) 10 10 $Line "#FFFFFF" $WeightPt | Out-Null
    }
}

function Add-VisioCopyMuxSymbolPx {
    param(
        [double]$X,
        [double]$Y,
        [double]$W,
        [double]$H,
        [string]$Line = "#000000",
        [double]$WeightPt = 1.4,
        [ValidateSet("right","left")][string]$Direction = "right"
    )
    if ($Direction -eq "right") {
        $pts = @(
            @($X,($Y + ($H * 0.10))),
            @(($X + ($W * 0.72)),($Y + ($H * 0.30))),
            @(($X + $W),($Y + ($H * 0.50))),
            @(($X + ($W * 0.72)),($Y + ($H * 0.70))),
            @($X,($Y + ($H * 0.90)))
        )
    } else {
        $pts = @(
            @(($X + $W),($Y + ($H * 0.10))),
            @(($X + ($W * 0.28)),($Y + ($H * 0.30))),
            @($X,($Y + ($H * 0.50))),
            @(($X + ($W * 0.28)),($Y + ($H * 0.70))),
            @(($X + $W),($Y + ($H * 0.90)))
        )
    }
    Add-VisioCopyPolylinePx -Points $pts -Line $Line -WeightPt $WeightPt -Closed | Out-Null
}

function Add-VisioCopyAndGateSymbolPx {
    param(
        [double]$X,
        [double]$Y,
        [double]$W,
        [double]$H,
        [string]$Line = "#000000",
        [double]$WeightPt = 1.6,
        [switch]$OutputBubble
    )
    $pts = @(
        @($X,$Y),
        @(($X + ($W * 0.56)),$Y),
        @(($X + ($W * 0.82)),($Y + ($H * 0.18))),
        @(($X + $W),($Y + ($H * 0.50))),
        @(($X + ($W * 0.82)),($Y + ($H * 0.82))),
        @(($X + ($W * 0.56)),($Y + $H)),
        @($X,($Y + $H))
    )
    Add-VisioCopyPolylinePx -Points $pts -Line $Line -WeightPt $WeightPt -Closed | Out-Null
    if ($OutputBubble) {
        Add-VisioCopyOvalPx ($X + $W - 2) ($Y + ($H / 2) - 6) 12 12 $Line "#FFFFFF" $WeightPt | Out-Null
    }
}

function Add-VisioCopyOrGateSymbolPx {
    param(
        [double]$X,
        [double]$Y,
        [double]$W,
        [double]$H,
        [string]$Line = "#000000",
        [double]$WeightPt = 1.6,
        [switch]$OutputBubble
    )
    $pts = @(
        @(($X + ($W * 0.08)),$Y),
        @(($X + ($W * 0.58)),($Y + ($H * 0.02))),
        @(($X + ($W * 0.88)),($Y + ($H * 0.22))),
        @(($X + $W),($Y + ($H * 0.50))),
        @(($X + ($W * 0.88)),($Y + ($H * 0.78))),
        @(($X + ($W * 0.58)),($Y + ($H * 0.98))),
        @(($X + ($W * 0.08)),($Y + $H)),
        @(($X + ($W * 0.28)),($Y + ($H * 0.50)))
    )
    Add-VisioCopyPolylinePx -Points $pts -Line $Line -WeightPt $WeightPt -Closed | Out-Null
    if ($OutputBubble) {
        Add-VisioCopyOvalPx ($X + $W - 2) ($Y + ($H / 2) - 7) 14 14 $Line "#FFFFFF" $WeightPt | Out-Null
    }
}

function Get-VisioCopyFitFontSize {
    param(
        [string]$Text,
        [double]$BoxWidthPx,
        [double]$RequestedPt,
        [double]$MinPt = 8,
        [string]$FontFace = "Arial"
    )
    $longest = 0
    foreach ($line in ($Text -split "`n")) {
        if ($line.Length -gt $longest) { $longest = $line.Length }
    }
    if ($longest -le 0) { return $RequestedPt }

    # Empirical guard for Visio export: monospaced fonts need a wider per-character estimate.
    $charFactor = 0.72
    if ($FontFace -match "Courier|Consolas|Mono") { $charFactor = 0.84 }
    $estimatedWidth = $longest * $RequestedPt * $charFactor
    if ($estimatedWidth -le ($BoxWidthPx * 0.94)) { return $RequestedPt }
    $fit = [Math]::Floor(($BoxWidthPx * 0.94) / ($longest * $charFactor))
    return [Math]::Max($MinPt, [Math]::Min($RequestedPt, $fit))
}

function Add-VisioCopyTextJobPx {
    param(
        [double]$X,
        [double]$Y,
        [double]$W,
        [double]$H,
        [string]$Text,
        [double]$SizePt = 14,
        [string]$Color = "#000000",
        [switch]$Bold,
        [switch]$Italic,
        [switch]$Underline,
        [string]$Align = "center",
        [double]$Angle = 0,
        [string]$FontFace = "Arial",
        [switch]$Fit
    )
    $drawX = $X
    $drawY = $Y
    $drawW = $W
    $drawH = $H
    $normAngle = [Math]::Abs([double]$Angle) % 180
    if ($normAngle -gt 80 -and $normAngle -lt 100) {
        $cx = $X + ($W / 2)
        $cy = $Y + ($H / 2)
        $drawW = $H
        $drawH = $W
        $drawX = $cx - ($drawW / 2)
        $drawY = $cy - ($drawH / 2)
    }
    if ($Fit) {
        $SizePt = Get-VisioCopyFitFontSize -Text $Text -BoxWidthPx $drawW -RequestedPt $SizePt -FontFace $FontFace
    }
    $script:VisioCopyTextJobs += [pscustomobject]@{
        X = $drawX; Y = $drawY; W = $drawW; H = $drawH; Text = $Text; SizePt = $SizePt
        Color = $Color; Bold = [bool]$Bold; Italic = [bool]$Italic; Underline = [bool]$Underline; Align = $Align; Angle = $Angle; FontFace = $FontFace
    }
}

function Add-VisioCopyTextNowPx {
    param($Job)
    $shape = Add-VisioCopyRectPx $Job.X $Job.Y $Job.W $Job.H "none" "none" 0
    $shape.Text = $Job.Text
    $fontFace = [string](Get-VisioCopyValue $Job "FontFace" "Arial")
    try { $shape.CellsU("Char.Font").FormulaU = "FONT(`"$fontFace`")" } catch {}
    try { $shape.CellsU("Char.Size").FormulaU = "$($Job.SizePt) pt" } catch {}
    try {
        $style = 0
        if ($Job.Bold) { $style += 1 }
        if ((Get-VisioCopyValue $Job "Italic" $false)) { $style += 2 }
        if ((Get-VisioCopyValue $Job "Underline" $false)) { $style += 4 }
        $shape.CellsU("Char.Style").FormulaU = "$style"
    } catch {}
    try { Set-VisioCopyColorCell $shape "Char.Color" $Job.Color } catch {}
    $ha = 1
    if ($Job.Align -eq "left") { $ha = 0 }
    if ($Job.Align -eq "right") { $ha = 2 }
    try {
        $shape.CellsU("Para.HorzAlign").FormulaU = "$ha"
        $shape.CellsU("VerticalAlign").FormulaU = "1"
        $shape.CellsU("LeftMargin").FormulaU = "0.01 in"
        $shape.CellsU("RightMargin").FormulaU = "0.01 in"
        $shape.CellsU("TopMargin").FormulaU = "0.01 in"
        $shape.CellsU("BottomMargin").FormulaU = "0.01 in"
        $shape.CellsU("Angle").FormulaU = "$($Job.Angle) deg"
        $shape.BringToFront() | Out-Null
    } catch {}
    return $shape
}

function Flush-VisioCopyText {
    foreach ($job in $script:VisioCopyTextJobs) {
        Add-VisioCopyTextNowPx $job | Out-Null
    }
    $script:VisioCopyTextJobs = @()
}

function Add-VisioCopyBoxLabelPx {
    param(
        [double]$X,
        [double]$Y,
        [double]$W,
        [double]$H,
        [string]$Text,
        [string]$Fill = "#FFFFFF",
        [double]$SizePt = 14,
        [string]$Line = "#000000",
        [double]$WeightPt = 0.8,
        [string]$TextColor = "#000000",
        [double]$Angle = 0,
        [double]$RoundPx = 0,
        [switch]$Bold,
        [switch]$Italic,
        [switch]$Underline,
        [switch]$Fit
    )
    Add-VisioCopyRectPx $X $Y $W $H $Line $Fill $WeightPt -RoundPx $RoundPx | Out-Null
    Add-VisioCopyTextJobPx $X $Y $W $H $Text $SizePt $TextColor -Bold:$Bold -Italic:$Italic -Underline:$Underline -Angle $Angle -Fit:$Fit
}

function Add-VisioCopyOrthogonalRoutePx {
    param(
        [Parameter(Mandatory=$true)][object[]]$Points,
        [string]$Color = "#000000",
        [double]$WeightPt = 1.0,
        [switch]$Dash,
        [int]$EndArrow = 13,
        [int]$BeginArrow = 0
    )
    for ($i = 0; $i -lt ($Points.Count - 1); $i++) {
        $p1 = $Points[$i]
        $p2 = $Points[$i + 1]
        $ea = 0
        $ba = 0
        if ($i -eq ($Points.Count - 2)) { $ea = $EndArrow }
        if ($i -eq 0) { $ba = $BeginArrow }
        Add-VisioCopyLinePx $p1[0] $p1[1] $p2[0] $p2[1] $Color $WeightPt -Dash:$Dash -EndArrow $ea -BeginArrow $ba | Out-Null
    }
}

function Add-VisioCopyTablePx {
    param(
        [double]$X,
        [double]$Y,
        [double]$W,
        [double]$H,
        [double[]]$ColumnWidths,
        [double[]]$RowHeights,
        [object[]]$Texts = @(),
        [string]$Line = "#000000",
        [string]$Fill = "#FFFFFF",
        [double]$WeightPt = 0.8,
        [double]$FontSizePt = 12
    )
    Add-VisioCopyRectPx $X $Y $W $H $Line $Fill $WeightPt | Out-Null
    $cx = $X
    for ($i = 0; $i -lt ($ColumnWidths.Count - 1); $i++) {
        $cx += $ColumnWidths[$i]
        Add-VisioCopyLinePx $cx $Y $cx ($Y + $H) $Line $WeightPt | Out-Null
    }
    $cy = $Y
    for ($i = 0; $i -lt ($RowHeights.Count - 1); $i++) {
        $cy += $RowHeights[$i]
        Add-VisioCopyLinePx $X $cy ($X + $W) $cy $Line $WeightPt | Out-Null
    }
    foreach ($cell in $Texts) {
        $col = [int](Get-VisioCopyValue $cell "col" 0)
        $row = [int](Get-VisioCopyValue $cell "row" 0)
        $tx = $X
        for ($i = 0; $i -lt $col; $i++) { $tx += $ColumnWidths[$i] }
        $ty = $Y
        for ($i = 0; $i -lt $row; $i++) { $ty += $RowHeights[$i] }
        Add-VisioCopyTextJobPx $tx $ty $ColumnWidths[$col] $RowHeights[$row] ([string](Get-VisioCopyValue $cell "text" "")) $FontSizePt "#000000" -Fit
    }
}

function Add-VisioCopyMessageLaneTablePx {
    param(
        [double]$X,
        [double]$Y,
        [double]$W,
        [double]$H,
        [object[]]$Rows,
        [double]$SideWidthPx = 40,
        [string]$LeftLabel = "Network",
        [string]$RightLabel = "Cache"
    )
    Add-VisioCopyRectPx $X $Y $W $H "#000000" "none" 0.8 -Dash | Out-Null
    Add-VisioCopyRectPx ($X + 10) ($Y + 8) $SideWidthPx ($H - 16) "#000000" "#D9D9D9" 0.7 | Out-Null
    Add-VisioCopyRectPx ($X + $W - 10 - $SideWidthPx) ($Y + 8) $SideWidthPx ($H - 16) "#000000" "#D9D9D9" 0.7 | Out-Null
    Add-VisioCopyTextJobPx ($X + 2) ($Y + ($H / 2) - 15) 110 26 $LeftLabel 12 "#000000" -Angle 90 -Fit
    Add-VisioCopyTextJobPx ($X + $W - 112) ($Y + ($H / 2) - 15) 110 26 $RightLabel 12 "#000000" -Angle 90 -Fit

    $laneX = $X + 10 + $SideWidthPx
    $laneW = $W - (2 * (10 + $SideWidthPx))
    $rowH = ($H - 16) / [Math]::Max(1, $Rows.Count)
    for ($i = 0; $i -lt $Rows.Count; $i++) {
        $row = $Rows[$i]
        $ry = $Y + 8 + ($i * $rowH)
        $fill = [string](Get-VisioCopyValue $row "fill" "#DDEED6")
        Add-VisioCopyRectPx $laneX $ry $laneW $rowH "#000000" $fill 0.6 | Out-Null
        Add-VisioCopyTextJobPx ($laneX + 16) $ry ($laneW - 32) $rowH ([string](Get-VisioCopyValue $row "text" "")) 12 "#000000" -Fit
        $direction = [string](Get-VisioCopyValue $row "direction" "right")
        if ($direction -eq "left") {
            Add-VisioCopyLinePx ($laneX + $laneW - 8) ($ry + ($rowH / 2)) ($laneX + 8) ($ry + ($rowH / 2)) "#000000" 1.1 -EndArrow 13 | Out-Null
        } else {
            Add-VisioCopyLinePx ($laneX + 8) ($ry + ($rowH / 2)) ($laneX + $laneW - 8) ($ry + ($rowH / 2)) "#000000" 1.1 -EndArrow 13 | Out-Null
        }
    }
}

function Add-VisioCopyHatchedRectPx {
    param(
        [double]$X,
        [double]$Y,
        [double]$W,
        [double]$H,
        [string]$Line = "#000000",
        [string]$Fill = "#FFFFFF",
        [string]$HatchColor = "#DADADA",
        [double]$StepPx = 10
    )
    Add-VisioCopyRectPx $X $Y $W $H $Line $Fill 0.8 | Out-Null
    for ($sx = $X - $H; $sx -lt ($X + $W); $sx += $StepPx) {
        $x1 = [Math]::Max($X, $sx)
        $y1 = $Y + ($x1 - $sx)
        $x2 = [Math]::Min($X + $W, $sx + $H)
        $y2 = $Y + ($x2 - $sx)
        Add-VisioCopyLinePx $x1 $y1 $x2 $y2 $HatchColor 0.45 | Out-Null
    }
}

function Add-VisioCopyBlockDownArrowPx {
    param(
        [double]$X,
        [double]$Y,
        [double]$W,
        [double]$H,
        [string]$Line = "#000000",
        [string]$Fill = "#FFFFFF",
        [double]$WeightPt = 1.6
    )
    $shaftW = $W * 0.34
    $shaftX = $X + (($W - $shaftW) / 2)
    $headY = $Y + ($H * 0.45)
    Add-VisioCopyRectPx $shaftX $Y $shaftW ($H * 0.48) $Line $Fill $WeightPt | Out-Null
    Add-VisioCopyLinePx $X $headY ($X + ($W / 2)) ($Y + $H) $Line $WeightPt | Out-Null
    Add-VisioCopyLinePx ($X + $W) $headY ($X + ($W / 2)) ($Y + $H) $Line $WeightPt | Out-Null
    Add-VisioCopyLinePx $X $headY $shaftX $headY $Line $WeightPt | Out-Null
    Add-VisioCopyLinePx ($shaftX + $shaftW) $headY ($X + $W) $headY $Line $WeightPt | Out-Null
}

function Get-VisioCopyIsoGridPoint {
    param(
        [double]$Ox,
        [double]$Oy,
        [int]$Row,
        [int]$Col,
        [double]$Dx = 66,
        [double]$Dy = 33,
        [double]$Skew = -33
    )
    return @(($Ox + ($Col * $Dx) + ($Row * $Skew)), ($Oy + ($Row * $Dy)))
}

function Add-VisioCopyIsoRouterGridPx {
    param(
        [double]$Ox,
        [double]$Oy,
        [int]$Rows = 4,
        [int]$Cols = 4,
        [int[]]$BoundaryIndexes = @(),
        [hashtable]$NodeFills = @{},
        [string]$ChipletFill = "#FFFFFF",
        [string]$BoundaryFill = "#CFCFCF",
        [string]$PlaneLine = "#BEBEBE",
        [double]$NodeRadiusPx = 12,
        [double]$Dx = 66,
        [double]$Dy = 33,
        [double]$Skew = -33,
        [double]$PlaneWidthPx = 230,
        [double]$PlaneDepthX = 44,
        [double]$PlaneDepthY = -31
    )
    $points = @{}
    for ($r = 0; $r -lt [Math]::Min(3, $Rows); $r++) {
        $p0 = Get-VisioCopyIsoGridPoint $Ox $Oy $r 0 $Dx $Dy $Skew
        Add-VisioCopyLinePx ($p0[0] - 28) ($p0[1] - 13) ($p0[0] + $PlaneWidthPx) ($p0[1] - 13) $PlaneLine 0.6 | Out-Null
        Add-VisioCopyLinePx ($p0[0] - 28) ($p0[1] - 13) ($p0[0] - 28 + $PlaneDepthX) ($p0[1] - 13 + $PlaneDepthY) $PlaneLine 0.6 | Out-Null
        Add-VisioCopyLinePx ($p0[0] + $PlaneWidthPx) ($p0[1] - 13) ($p0[0] + $PlaneWidthPx + $PlaneDepthX) ($p0[1] - 13 + $PlaneDepthY) $PlaneLine 0.6 | Out-Null
        Add-VisioCopyLinePx ($p0[0] - 28 + $PlaneDepthX) ($p0[1] - 13 + $PlaneDepthY) ($p0[0] + $PlaneWidthPx + $PlaneDepthX) ($p0[1] - 13 + $PlaneDepthY) $PlaneLine 0.6 | Out-Null
    }

    for ($r = 0; $r -lt $Rows; $r++) {
        for ($c = 0; $c -lt $Cols; $c++) {
            $points["$r,$c"] = Get-VisioCopyIsoGridPoint $Ox $Oy $r $c $Dx $Dy $Skew
        }
    }
    for ($r = 0; $r -lt $Rows; $r++) {
        for ($c = 0; $c -lt ($Cols - 1); $c++) {
            $a = $points["$r,$c"]
            $b = $points["$r,$($c + 1)"]
            Add-VisioCopyLinePx $a[0] $a[1] $b[0] $b[1] "#000000" 0.8 | Out-Null
        }
    }
    for ($c = 0; $c -lt $Cols; $c++) {
        for ($r = 0; $r -lt ($Rows - 1); $r++) {
            $a = $points["$r,$c"]
            $b = $points["$($r + 1),$c"]
            Add-VisioCopyLinePx $a[0] $a[1] $b[0] $b[1] "#000000" 0.8 | Out-Null
        }
    }
    $idx = 0
    for ($r = 0; $r -lt $Rows; $r++) {
        for ($c = 0; $c -lt $Cols; $c++) {
            $p = $points["$r,$c"]
            $fill = $ChipletFill
            if ($BoundaryIndexes -contains $idx) { $fill = $BoundaryFill }
            if ($NodeFills.ContainsKey("$r,$c")) { $fill = [string]$NodeFills["$r,$c"] }
            Add-VisioCopyOvalPx ($p[0] - $NodeRadiusPx) ($p[1] - $NodeRadiusPx) ($NodeRadiusPx * 2) ($NodeRadiusPx * 2) "#000000" $fill 0.8 | Out-Null
            $idx++
        }
    }
    return $points
}
