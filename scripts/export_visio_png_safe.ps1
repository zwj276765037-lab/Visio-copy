param(
    [Parameter(Mandatory=$true)][string]$VsdxPath,
    [string]$PngPath,
    [int]$TimeoutSeconds = 45
)

$ErrorActionPreference = "Stop"

$resolvedVsdx = (Resolve-Path -LiteralPath $VsdxPath).Path
if ([string]::IsNullOrWhiteSpace($PngPath)) {
    $PngPath = [System.IO.Path]::ChangeExtension($resolvedVsdx, ".png")
}
$pngFull = [System.IO.Path]::GetFullPath($PngPath)
if (Test-Path -LiteralPath $pngFull) {
    Remove-Item -LiteralPath $pngFull -Force
}

$start = Get-Date
$job = Start-Job -ArgumentList $resolvedVsdx,$pngFull -ScriptBlock {
    param($Vsdx,$Png)
    $ErrorActionPreference = "Stop"
    $visio = $null
    $doc = $null
    try {
        $visio = New-Object -ComObject Visio.Application
        $visio.Visible = $false
        try { $visio.AlertResponse = 7 } catch {}
        try { $visio.DisplayAlerts = 0 } catch {}
        $openFlags = 2 + 64 + 128
        $doc = $visio.Documents.OpenEx($Vsdx, $openFlags)
        $doc.Pages.Item(1).Export($Png)
        [pscustomobject]@{
            ok = (Test-Path -LiteralPath $Png)
            png = $Png
            bytes = (Get-Item -LiteralPath $Png -ErrorAction SilentlyContinue).Length
        }
    } finally {
        if ($null -ne $doc) { try { $doc.Close() } catch {} }
        if ($null -ne $visio) { try { $visio.Quit() } catch {} }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

$done = Wait-Job $job -Timeout $TimeoutSeconds
if ($done) {
    $result = Receive-Job $job
    Remove-Job $job
    if (-not (Test-Path -LiteralPath $pngFull)) {
        throw "Visio export reported completion but PNG is missing: $pngFull"
    }
    $result
} else {
    Stop-Job $job -Force
    Remove-Job $job -Force
    Get-CimInstance Win32_Process -Filter "name='VISIO.EXE'" |
        Where-Object { $_.CreationDate -gt $start } |
        ForEach-Object { Stop-Process -Id $_.ProcessId -Force }
    throw "Visio Page.Export timed out after $TimeoutSeconds seconds: $resolvedVsdx"
}
