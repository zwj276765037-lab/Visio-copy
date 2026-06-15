param(
    [Parameter(Mandatory=$true)][string]$Svg,
    [Parameter(Mandatory=$true)][string]$OutVsdx,
    [Parameter(Mandatory=$true)][int]$SourceWidthPx,
    [Parameter(Mandatory=$true)][int]$SourceHeightPx,
    [bool]$Visible = $true,
    [bool]$KeepOpen = $true
)

$ErrorActionPreference = "Stop"

throw "Disabled: visio-copy must draw with native Visio shapes only; SVG/PDF/vector trace import is forbidden."
