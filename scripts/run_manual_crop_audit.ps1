param(
    [string]$OutputDir = $PSScriptRoot,
    [int]$Zoom = 3,
    [string]$PythonExe = "",
    [double]$WarnThreshold = 0.15,
    [string]$SummaryPath = ""
)

$ErrorActionPreference = "Stop"

$ResolvedOutputDir = (Resolve-Path -LiteralPath $OutputDir).Path
$ManifestPath = Join-Path $ResolvedOutputDir "manifest.json"
if (-not (Test-Path -LiteralPath $ManifestPath)) {
    throw "manifest.json not found: $ManifestPath"
}

$Manifest = Get-Content -LiteralPath $ManifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
if ($null -eq $Manifest.audit_components -or @($Manifest.audit_components).Count -eq 0) {
    throw "manifest.json has no audit_components array: $ManifestPath"
}

$SourcePath = Join-Path $ResolvedOutputDir ([string]$Manifest.source)
$PreviewPath = Join-Path $ResolvedOutputDir ([string]$Manifest.preview)
if (-not (Test-Path -LiteralPath $SourcePath)) {
    throw "source image not found: $SourcePath"
}
if (-not (Test-Path -LiteralPath $PreviewPath)) {
    throw "preview image not found: $PreviewPath"
}

if ([string]::IsNullOrWhiteSpace($PythonExe)) {
    $Bundled = Join-Path $env:USERPROFILE ".cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
    if (Test-Path -LiteralPath $Bundled) {
        $PythonExe = $Bundled
    } else {
        $PythonExe = "python"
    }
}

$SkillDir = Join-Path $env:USERPROFILE ".codex\skills\visio-copy"
$CropCompare = Join-Path $SkillDir "scripts\crop_compare.py"
$AuditDir = Join-Path $ResolvedOutputDir "audit"

$Args = @(
    $CropCompare,
    $SourcePath,
    $PreviewPath,
    "--out",
    $AuditDir,
    "--zoom",
    ([string]$Zoom)
)

foreach ($Component in @($Manifest.audit_components)) {
    $Name = [string]$Component.name
    $X = [int]$Component.x
    $Y = [int]$Component.y
    $W = [int]$Component.w
    $H = [int]$Component.h
    if ([string]::IsNullOrWhiteSpace($Name) -or $W -le 0 -or $H -le 0) {
        throw "invalid audit component in manifest: $($Component | ConvertTo-Json -Compress)"
    }
    $Args += "--component"
    $Args += "$($Name):$X,$Y,$W,$H"
}

& $PythonExe @Args
if ($LASTEXITCODE -ne 0) {
    throw "crop_compare.py failed with exit code $LASTEXITCODE"
}

$Rows = Get-ChildItem -LiteralPath $AuditDir -Filter "*.metrics.json" |
    ForEach-Object {
        $Metrics = Get-Content -LiteralPath $_.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
        [pscustomobject]@{
            name = $Metrics.name
            changed_pixel_fraction = $Metrics.changed_pixel_fraction
            metrics = $_.Name
        }
    } |
    Sort-Object changed_pixel_fraction -Descending

if ([string]::IsNullOrWhiteSpace($SummaryPath)) {
    $SummaryPath = Join-Path $ResolvedOutputDir "audit_summary.md"
}

$Summary = New-Object System.Collections.Generic.List[string]
$Summary.Add("# Audit Summary")
$Summary.Add("")
$Summary.Add('Generated with `run_manual_crop_audit.ps1` from `manifest.json` `audit_components`.')
$Summary.Add("")
$Summary.Add("| Component | Changed pixel fraction | Status |")
$Summary.Add("| --- | ---: | --- |")
foreach ($Row in $Rows) {
    $status = "ok"
    if ([double]$Row.changed_pixel_fraction -gt $WarnThreshold) { $status = "review" }
    $Summary.Add(("| {0} | {1:N6} | {2} |" -f $Row.name, [double]$Row.changed_pixel_fraction, $status))
}
Set-Content -LiteralPath $SummaryPath -Value $Summary -Encoding UTF8

$HighRows = @($Rows | Where-Object { [double]$_.changed_pixel_fraction -gt $WarnThreshold })
if ($HighRows.Count -gt 0) {
    Write-Warning ("{0} component(s) exceed changed_pixel_fraction threshold {1:N3}; inspect audit_summary.md and side-by-side crops before accepting." -f $HighRows.Count, $WarnThreshold)
}

$Rows
