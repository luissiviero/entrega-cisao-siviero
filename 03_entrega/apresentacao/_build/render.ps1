# Exports every slide of a deck in 03_entrega\apresentacao as PNG, using the
# installed PowerPoint (read-only open, no window). Output:
#   03_entrega\apresentacao\_build\png\<deck name>\sNN.png
# Usage: powershell.exe -NoProfile -File <this script> <deck.pptx>
param([Parameter(Mandatory = $true)][string]$Deck)
$ErrorActionPreference = "Stop"

$root = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$deckPath = (Resolve-Path $Deck).Path
if (-not $deckPath.StartsWith($root, [StringComparison]::OrdinalIgnoreCase) -or
    [IO.Path]::GetExtension($deckPath) -ne ".pptx") {
    throw "Only .pptx files inside $root are accepted: $deckPath"
}

$outDir = Join-Path $PSScriptRoot ("png\" + [IO.Path]::GetFileNameWithoutExtension($deckPath))
New-Item -ItemType Directory -Force -Path $outDir | Out-Null
Get-ChildItem $outDir -Filter "s*.png" | Remove-Item

# PowerPoint is single-instance: if the user already has it open, leave it running.
$wasRunning = [bool](Get-Process POWERPNT -ErrorAction SilentlyContinue)
$app = New-Object -ComObject PowerPoint.Application
try {
    # Open(FileName, ReadOnly = msoTrue, Untitled = msoFalse, WithWindow = msoFalse)
    $pres = $app.Presentations.Open($deckPath, -1, 0, 0)
    try {
        $i = 1
        foreach ($slide in $pres.Slides) {
            $slide.Export((Join-Path $outDir ("s{0:D2}.png" -f $i)), "PNG", 1600, 900)
            $i++
        }
        "exported $($i - 1) slides to $outDir"
    } finally {
        $pres.Close()
    }
} finally {
    if (-not $wasRunning) { $app.Quit() }
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)
}
