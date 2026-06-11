$ErrorActionPreference = 'Continue'
$dest = "C:\geminiterminal2\3800_NATIONWIDE_Pleadings_PDFs"
$manifest = Import-Csv (Join-Path $dest "_manifest.csv")
$GUID = "7c5232e4-d263-459d-95f2-0b709699a656"
$cacheRoot = "$env:LOCALAPPDATA\Egnyte Connect\data\$GUID"
$MIN_FREE_GB = 6.0          # when free drops below this, pause & clear Egnyte cache
$SAFE_AFTER_GB = 6.0        # if still below this right after a clear, we're truly out of room

function FreeGB { (Get-PSDrive C).Free/1GB }

function Clear-And-Remount {
    Write-Host ">>> Clearing Egnyte cache (free=$([math]::Round((FreeGB),1))GB)..."
    Get-Process EgnyteClient,EgnyteDrive,EgnyteSyncService -ErrorAction SilentlyContinue | Stop-Process -Force
    Start-Sleep 4
    if (Test-Path "$cacheRoot\data") { Remove-Item "$cacheRoot\data" -Recurse -Force -ErrorAction SilentlyContinue }
    if (Test-Path "$cacheRoot\cacheDB.sqlite") { Remove-Item "$cacheRoot\cacheDB.sqlite" -Force -ErrorAction SilentlyContinue }
    Start-Sleep 2
    $exe = "C:\Program Files\Egnyte Connect\EgnyteClient.exe"
    schtasks /Delete /TN "TmpStartEgnyte" /F 2>$null | Out-Null
    schtasks /Create /TN "TmpStartEgnyte" /TR "`"$exe`"" /SC ONCE /ST 23:59 /RL LIMITED /F /IT 2>&1 | Out-Null
    schtasks /Run /TN "TmpStartEgnyte" 2>&1 | Out-Null
    for ($k=0; $k -lt 40; $k++) { if (Test-Path '\\EgnyteDrive\bordinsemmer') { break }; Start-Sleep 3 }
    schtasks /Delete /TN "TmpStartEgnyte" /F 2>&1 | Out-Null
    Write-Host ">>> Remounted. Z online=$(Test-Path 'Z:\Shared\Current Clients\3800- NATIONWIDE') free=$([math]::Round((FreeGB),1))GB"
}

$ok=0; $skip=0; $fail=0; $i=0; $tot=$manifest.Count
foreach ($row in $manifest) {
    $i++
    $target = Join-Path $dest $row.DestFile
    if (Test-Path -LiteralPath $target) { $skip++; continue }
    if ((FreeGB) -lt $MIN_FREE_GB) {
        Clear-And-Remount
        if ((FreeGB) -lt $SAFE_AFTER_GB) {
            Write-Host "=== OUT OF SPACE after clear (free=$([math]::Round((FreeGB),1))GB) at [$i/$tot]. copied=$ok skip=$skip. STOPPING. ==="
            break
        }
    }
    if (-not (Test-Path -LiteralPath $row.SourcePath)) { Write-Host "MISSING SRC: $($row.SourcePath)"; $fail++; continue }
    try { Copy-Item -LiteralPath $row.SourcePath -Destination $target -ErrorAction Stop; $ok++ }
    catch { Write-Host "FAILED: $($row.SourcePath) -> $($_.Exception.Message)"; $fail++ }
    if ($i % 500 -eq 0) { Write-Host "[copy $i/$tot] copied=$ok skip=$skip fail=$fail free=$([math]::Round((FreeGB),1))GB" }
}
$onDisk = (Get-ChildItem -LiteralPath $dest -File -Filter *.pdf).Count
$remaining = ($manifest | Where-Object { -not (Test-Path -LiteralPath (Join-Path $dest $_.DestFile)) }).Count
Write-Host "=== DONE. planned=$tot copied-this-run=$ok skip(existing)=$skip fail=$fail on-disk=$onDisk remaining=$remaining free=$([math]::Round((FreeGB),1))GB ==="
