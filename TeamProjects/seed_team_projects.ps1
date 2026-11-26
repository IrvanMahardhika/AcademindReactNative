# Creates dummy payloads so TreeSize MB figures match the gold report.
$Base = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "Team_Projects"
New-Item -ItemType Directory -Force -Path $Base | Out-Null

$MiB = 1MB
$folders = [ordered]@{
    "AtlasCore"   = [int64](48 * $MiB)
    "BeaconDocs"  = [int64](12.5 * $MiB)
    "CedarMedia"  = [int64](36 * $MiB)
    "DriftLogs"   = [int64](8 * $MiB)
    "EmberData"   = [int64](64 * $MiB)
    "FalconLegal" = [int64](15 * $MiB)
    "HarborML"    = [int64](22 * $MiB)
    "IrisWeb"     = [int64](10.5 * $MiB)
}

foreach ($name in $folders.Keys) {
    $dir = Join-Path $Base $name
    New-Item -ItemType Directory -Force -Path $dir | Out-Null
    $payload = Join-Path $dir "payload.bin"
    $size = $folders[$name]
    if (Test-Path $payload) { Remove-Item $payload -Force }
    fsutil file createnew $payload $size | Out-Null
    Write-Host "$name $size"
}
Write-Host "Seeded $Base"
