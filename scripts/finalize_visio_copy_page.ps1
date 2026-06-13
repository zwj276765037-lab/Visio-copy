param(
    [Parameter(Mandatory = $true)]
    [string]$TargetPath,

    [Parameter(Mandatory = $true)]
    [string]$PageName,

    [Parameter(Mandatory = $true)]
    [string]$FinalLayerName,

    [string]$BackupSuffix = "before-visio-copy-final-cleanup"
)

$ErrorActionPreference = "Stop"

function Set-CellFormula {
    param($Shape, [string]$Cell, [string]$Formula)
    try { $Shape.CellsU($Cell).FormulaU = $Formula } catch {}
}

function Test-ShapeOnLayer {
    param($Shape, [string]$LayerName)
    try {
        if ($Shape.LayerCount -le 0) { return $false }
        for ($i = 1; $i -le $Shape.LayerCount; $i++) {
            if ($Shape.Layer($i).NameU -eq $LayerName -or $Shape.Layer($i).Name -eq $LayerName) {
                return $true
            }
        }
    } catch {}
    return $false
}

function Remove-EmptyNonFinalLayers {
    param($Page, [string]$FinalLayerName)
    $layers = @()
    foreach ($layer in $Page.Layers) { $layers += $layer }
    foreach ($layer in $layers) {
        try {
            if ($layer.NameU -ne $FinalLayerName -and $layer.Name -ne $FinalLayerName) {
                $layer.Delete(1)
            }
        } catch {}
    }
}

$TargetPath = [System.IO.Path]::GetFullPath($TargetPath)
if (-not (Test-Path -LiteralPath $TargetPath)) {
    throw "Target .vsdx not found: $TargetPath"
}

$targetDir = [System.IO.Path]::GetDirectoryName($TargetPath)
$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$backup = Join-Path $targetDir ("hardware-diagram.{0}.{1}.vsdx" -f $BackupSuffix, $timestamp)
Copy-Item -LiteralPath $TargetPath -Destination $backup -Force

try {
    $visio = [Runtime.InteropServices.Marshal]::GetActiveObject("Visio.Application")
} catch {
    $visio = New-Object -ComObject Visio.Application
}
$visio.Visible = $true
try { $visio.UserControl = $true } catch {}

$doc = $null
foreach ($d in $visio.Documents) {
    if ([string]::Equals($d.FullName, $TargetPath, [System.StringComparison]::OrdinalIgnoreCase)) {
        $doc = $d
        break
    }
}
if ($null -eq $doc) { $doc = $visio.Documents.Open($TargetPath) }

$page = $null
try { $page = $doc.Pages.ItemU($PageName) } catch {
    try { $page = $doc.Pages.Item($PageName) } catch {}
}
if ($null -eq $page) {
    throw "Page not found: $PageName"
}

try { $null = $page.Layers.ItemU($FinalLayerName) } catch {
    throw "Final layer not found on page '$PageName': $FinalLayerName"
}

$toDelete = @()
foreach ($shape in $page.Shapes) {
    if (-not (Test-ShapeOnLayer $shape $FinalLayerName)) {
        $toDelete += $shape
    } else {
        foreach ($cell in @("LockMove","LockWidth","LockHeight","LockDelete","LockRotate","LockAspect")) {
            Set-CellFormula $shape $cell "0"
        }
    }
}

foreach ($shape in $toDelete) {
    try { $shape.Delete() } catch {}
}

Remove-EmptyNonFinalLayers $page $FinalLayerName

$doc.Save()

[pscustomobject]@{
    PageName = $PageName
    FinalLayerName = $FinalLayerName
    DeletedShapes = $toDelete.Count
    Backup = $backup
    Result = "Final page cleanup complete. Only shapes on the final layer were kept."
} | ConvertTo-Json
