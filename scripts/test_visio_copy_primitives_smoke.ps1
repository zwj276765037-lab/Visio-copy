param(
    [string]$OutputDir = (Join-Path $env:TEMP "visio-copy-helper-smoke"),
    [switch]$KeepArtifacts
)

$ErrorActionPreference = "Stop"

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
. (Join-Path $ScriptDir "visio_copy_manual_primitives.ps1")

New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null
$vsdx = Join-Path $OutputDir "primitive_smoke.vsdx"
$png = Join-Path $OutputDir "primitive_smoke.png"
if (Test-Path -LiteralPath $vsdx) { Remove-Item -LiteralPath $vsdx -Force }
if (Test-Path -LiteralPath $png) { Remove-Item -LiteralPath $png -Force }

$visio = $null
$doc = $null
try {
    $visio = New-Object -ComObject Visio.Application
    $visio.Visible = $false
    try { $visio.AlertResponse = 7 } catch {}
    try { $visio.DisplayAlerts = 0 } catch {}

    $doc = $visio.Documents.Add("")
    $page = $visio.ActivePage
    Initialize-VisioCopyCanvas -Page $page -SourceWidthPx 640 -SourceHeightPx 360

    Add-VisioCopyBoxLabelPx 20 20 150 44 "Long Controller Label" "#DDEAF7" 16 -Fit
    Add-VisioCopyOrthogonalRoutePx -Points @(@(50,90),@(180,90),@(180,140)) -Color "#000000" -EndArrow 13
    Add-VisioCopyTablePx -X 220 -Y 20 -W 220 -H 90 `
        -ColumnWidths @(70,80,70) `
        -RowHeights @(30,30,30) `
        -Texts @(
            @{ row = 0; col = 0; text = "Vnet" },
            @{ row = 0; col = 1; text = "Type" },
            @{ row = 0; col = 2; text = "Flit" }
        )
    Add-VisioCopyMessageLaneTablePx -X 20 -Y 160 -W 270 -H 120 `
        -Rows @(
            @{ text = "VN0 Request"; fill = "#DDEED6"; direction = "right" },
            @{ text = "VN1 Response"; fill = "#DDEED6"; direction = "left" },
            @{ text = "VN2 Request"; fill = "#B7D99E"; direction = "right" }
        )
    Add-VisioCopyHatchedRectPx 320 160 120 45
    Add-VisioCopyBlockDownArrowPx 470 20 70 80
    $points = Add-VisioCopyIsoRouterGridPx -Ox 480 -Oy 180 -BoundaryIndexes @(1,2,6)
    Add-VisioCopyLinePx $points["0,0"][0] $points["0,0"][1] $points["3,3"][0] $points["3,3"][1] "#000000" 0.8 -Dash | Out-Null

    Flush-VisioCopyText
    $doc.SaveAs($vsdx)
} finally {
    if ($null -ne $doc) { try { $doc.Close() } catch {} }
    if ($null -ne $visio) { try { $visio.Quit() } catch {} }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

$exporter = Join-Path $ScriptDir "export_visio_png_safe.ps1"
& powershell -NoProfile -ExecutionPolicy Bypass -File $exporter -VsdxPath $vsdx -PngPath $png -TimeoutSeconds 45 | Out-Host

$result = [pscustomobject]@{
    vsdx = (Test-Path -LiteralPath $vsdx)
    png = (Test-Path -LiteralPath $png)
    pngBytes = (Get-Item -LiteralPath $png -ErrorAction SilentlyContinue).Length
    outputDir = $OutputDir
}
$result

if (-not $KeepArtifacts) {
    # Keep this conservative: only remove files created by this smoke test.
    foreach ($path in @($vsdx, $png)) {
        if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force }
    }
}
