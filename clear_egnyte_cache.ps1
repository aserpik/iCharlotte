$ErrorActionPreference = 'SilentlyContinue'
$sentinel = 'C:\geminiterminal2\logs\egnyte_clear_done.txt'
Remove-Item $sentinel -Force -ErrorAction SilentlyContinue

# 1. Stop the admin-protected update service so it stops respawning Egnyte processes
Stop-Service 'EgnyteConnectDesktopUpdate64bit' -Force -ErrorAction SilentlyContinue

# 2. Kill all Egnyte processes (retry loop to beat any respawn)
for ($i = 0; $i -lt 6; $i++) {
    Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.Name -like '*egnyte*' } | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Milliseconds 700
}
Start-Sleep -Seconds 2
$still = ((Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.Name -like '*egnyte*' }).Name | Sort-Object -Unique) -join ','

# 3. Clear ONLY the cache blocks + cache DB (keep local_data_db.sqlite and upload/)
$root      = 'C:\Users\ASerpik.DESKTOP-MRIMK0D\AppData\Local\Egnyte Connect\data'
$guidRoot  = Join-Path $root '7c5232e4-d263-459d-95f2-0b709699a656'
$cacheData = Join-Path $guidRoot 'data'
$tmp       = Join-Path $root 'tmp'

if (Test-Path $cacheData) { Get-ChildItem $cacheData -Force -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue }
Remove-Item (Join-Path $guidRoot 'cacheDB.sqlite')  -Force -ErrorAction SilentlyContinue
Remove-Item (Join-Path $guidRoot 'logFile.dat')     -Force -ErrorAction SilentlyContinue
Remove-Item (Join-Path $guidRoot 'tmpLogFile.dat')  -Force -ErrorAction SilentlyContinue
if (Test-Path $tmp) { Get-ChildItem $tmp -Force -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue }

# 4. Report
$free  = [math]::Round((Get-PSDrive C).Free / 1GB, 2)
$remain = 0
if (Test-Path $cacheData) { $remain = [math]::Round(((Get-ChildItem $cacheData -Recurse -File -Force -ErrorAction SilentlyContinue | Measure-Object Length -Sum).Sum)/1GB, 2) }
Set-Content $sentinel "DONE`r`nStillRunning: $still`r`nCacheRemainingGB: $remain`r`nFreeGB: $free"
